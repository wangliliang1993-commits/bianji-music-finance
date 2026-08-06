import { MovementType, OrderStatus, Prisma, TransactionType } from "@prisma/client";

const asNumber = (value: Prisma.Decimal) => Number(value.toString());

/**
 * 确认采购单：在一个事务里完成库存入库、移动平均成本更新、资金支出和审计。
 * 该函数应在 API 的管理员/成员权限校验之后调用。
 */
export async function confirmPurchase(
  db: Prisma.TransactionClient,
  purchaseId: string,
  actorId: string
) {
  const order = await db.purchaseOrder.findUnique({
    where: { id: purchaseId },
    include: { items: true }
  });
  if (!order) throw new Error("采购单不存在");
  if (order.status !== OrderStatus.DRAFT) throw new Error("只有草稿采购单可以确认");

  for (const item of order.items) {
    const current = await db.inventoryBalance.upsert({
      where: { productId: item.productId },
      create: { productId: item.productId, quantity: 0, averageCostFen: 0 },
      update: {}
    });
    const oldQty = asNumber(current.quantity);
    const inQty = asNumber(item.quantity);
    const newQty = oldQty + inQty;
    const newAverage = Math.round(
      (oldQty * Number(current.averageCostFen) + inQty * Number(item.unitCostFen)) / newQty
    );
    await db.inventoryBalance.update({
      where: { id: current.id },
      data: { quantity: new Prisma.Decimal(newQty), averageCostFen: BigInt(newAverage) }
    });
    await db.inventoryMovement.create({
      data: {
        productId: item.productId,
        type: MovementType.PURCHASE,
        quantity: item.quantity,
        unitCostFen: item.unitCostFen,
        referenceId: order.id
      }
    });
  }

  const transaction = await db.transaction.create({
    data: {
      type: TransactionType.EXPENSE,
      amountFen: order.paidFen,
      occurredAt: order.occurredAt,
      summary: `采购付款 ${order.orderNo}`,
      accountId: order.accountId,
      createdById: actorId
    }
  });
  const confirmed = await db.purchaseOrder.update({
    where: { id: order.id },
    data: { status: OrderStatus.CONFIRMED, transactionId: transaction.id }
  });
  await db.auditLog.create({
    data: { actorId, action: "CONFIRM", entityType: "PurchaseOrder", entityId: order.id, before: { status: "DRAFT" }, after: { status: "CONFIRMED" } }
  });
  return confirmed;
}

/** 确认销售单：先锁定业务结果，逐项检查库存，再结转成本并生成现金收入。 */
export async function confirmSale(
  db: Prisma.TransactionClient,
  saleId: string,
  actorId: string
) {
  const order = await db.salesOrder.findUnique({ where: { id: saleId }, include: { items: true } });
  if (!order) throw new Error("销售单不存在");
  if (order.status !== OrderStatus.DRAFT) throw new Error("只有草稿销售单可以确认");

  let totalCostFen = BigInt(0);
  for (const item of order.items) {
    const stock = await db.inventoryBalance.findUnique({
      where: { productId: item.productId }
    });
    const saleQty = asNumber(item.quantity);
    if (!stock || asNumber(stock.quantity) < saleQty) throw new Error(`商品 ${item.productId} 库存不足`);
    const unitCostFen = stock.averageCostFen;
    totalCostFen += BigInt(Math.round(saleQty * Number(unitCostFen)));
    await db.inventoryBalance.update({
      where: { id: stock.id },
      data: { quantity: new Prisma.Decimal(asNumber(stock.quantity) - saleQty) }
    });
    await db.salesItem.update({ where: { id: item.id }, data: { unitCostFen } });
    await db.inventoryMovement.create({
      data: {
        productId: item.productId,
        type: MovementType.SALE,
        quantity: new Prisma.Decimal(-saleQty),
        unitCostFen,
        referenceId: order.id
      }
    });
  }

  const transaction = await db.transaction.create({
    data: {
      type: TransactionType.INCOME,
      amountFen: order.receivedFen,
      occurredAt: order.occurredAt,
      summary: `商品销售 ${order.orderNo}`,
      accountId: order.accountId,
      createdById: actorId
    }
  });
  const confirmed = await db.salesOrder.update({
    where: { id: order.id },
    data: { status: OrderStatus.CONFIRMED, costFen: totalCostFen, transactionId: transaction.id }
  });
  await db.auditLog.create({
    data: { actorId, action: "CONFIRM", entityType: "SalesOrder", entityId: order.id, before: { status: "DRAFT" }, after: { status: "CONFIRMED", costFen: totalCostFen.toString() } }
  });
  return confirmed;
}

export const formatMoney = (fen: number, compact = false) => {
  if (compact && Math.abs(fen) >= 100_000_000) return `¥${(fen / 100_000_000).toFixed(1)}万`;
  return new Intl.NumberFormat("zh-CN", {
    style: "currency",
    currency: "CNY",
    minimumFractionDigits: 2
  }).format(fen / 100);
};

export type InventoryState = { quantity: number; averageCostFen: number };

export function receiveStock(state: InventoryState, quantity: number, unitCostFen: number): InventoryState {
  if (quantity <= 0 || unitCostFen < 0) throw new Error("采购数量和成本无效");
  const totalQuantity = state.quantity + quantity;
  const totalValue = state.quantity * state.averageCostFen + quantity * unitCostFen;
  return { quantity: totalQuantity, averageCostFen: Math.round(totalValue / totalQuantity) };
}

export function issueStock(state: InventoryState, quantity: number) {
  if (quantity <= 0) throw new Error("销售数量无效");
  if (quantity > state.quantity) throw new Error("库存不足");
  return {
    state: { quantity: state.quantity - quantity, averageCostFen: state.averageCostFen },
    costFen: quantity * state.averageCostFen
  };
}

export function grossProfitFen(receivedFen: number, costFen: number) {
  return receivedFen - costFen;
}

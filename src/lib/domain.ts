import { z } from "zod";

export const transactionInput = z.object({
  type: z.enum(["INCOME", "EXPENSE"]),
  amountFen: z.number().int().positive(),
  occurredAt: z.coerce.date(),
  summary: z.string().trim().min(1).max(100),
  accountId: z.string().min(1),
  categoryId: z.string().min(1),
  courseCategoryId: z.string().optional(),
  note: z.string().max(500).optional()
});

export const salesOrderInput = z.object({
  accountId: z.string().min(1),
  receivedFen: z.number().int().nonnegative(),
  discountFen: z.number().int().nonnegative(),
  occurredAt: z.coerce.date(),
  items: z.array(z.object({
    productId: z.string().min(1),
    quantity: z.number().positive(),
    unitPriceFen: z.number().int().nonnegative()
  })).min(1)
});

export const purchaseOrderInput = z.object({
  supplier: z.string().trim().min(1).max(100),
  accountId: z.string().min(1),
  feeFen: z.number().int().nonnegative(),
  occurredAt: z.coerce.date(),
  items: z.array(z.object({
    productId: z.string().min(1),
    quantity: z.number().positive(),
    unitCostFen: z.number().int().nonnegative()
  })).min(1)
});

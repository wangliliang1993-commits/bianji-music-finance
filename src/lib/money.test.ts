import { describe, expect, it } from "vitest";
import { grossProfitFen, issueStock, receiveStock } from "./money";

describe("移动加权平均库存", () => {
  it("连续采购后正确计算平均成本", () => {
    let stock = receiveStock({ quantity: 0, averageCostFen: 0 }, 10, 10000);
    stock = receiveStock(stock, 10, 20000);
    expect(stock).toEqual({ quantity: 20, averageCostFen: 15000 });
  });

  it("销售按当前平均成本结转", () => {
    const result = issueStock({ quantity: 20, averageCostFen: 15000 }, 3);
    expect(result.costFen).toBe(45000);
    expect(result.state.quantity).toBe(17);
    expect(grossProfitFen(60000, result.costFen)).toBe(15000);
  });

  it("禁止超库存销售", () => {
    expect(() => issueStock({ quantity: 2, averageCostFen: 100 }, 3)).toThrow("库存不足");
  });
});

# 边际音乐财务

边际音乐单店财务与进销存管理站点。当前版本包含经营概览、资金流水、课时费收入、采购销售、商品库存、移动加权平均成本、报表和基础设置界面，并提供完整 PostgreSQL 数据模型。

## 本地运行

```bash
npm install
cp .env.example .env
npm run dev
```

访问 `http://localhost:3000`。界面自带示例数据，方便直接评审产品体验。

## 数据库

在 `.env` 中配置 PostgreSQL 的 `DATABASE_URL`，然后执行：

```bash
npx prisma generate
npx prisma migrate dev --name init
```

金额统一以人民币分为单位保存，库存数量使用三位定点小数。销售成本采用移动加权平均法。所有采购、销售、调拨与退货的确认逻辑应放在单个数据库事务中完成。

## 验证

```bash
npm test
npm run build
```

当前测试覆盖移动加权平均成本、销售成本结转、毛利与超库存拦截。

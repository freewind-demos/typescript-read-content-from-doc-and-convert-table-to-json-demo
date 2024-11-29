# Word文档表格转JSON工具

这是一个使用TypeScript编写的命令行工具，可以从Word文档中提取指定标题后的表格，并将其转换为JSON格式。

## 功能特点

- 支持读取.docx格式的Word文档
- 可以通过标题定位目标表格
- 自动将表格内容转换为JSON格式
- 支持表格格式：
  - 第一行作为标题行（会被忽略）
  - 后续行的第一列作为key，第二列作为value

## 安装

```bash
# 克隆项目后，安装依赖
pnpm install
```

## 使用方法

1. 准备Word文档
   - 文档中应包含标题（如"系统配置"）
   - 标题后紧跟一个两列表格
   - 表格第一行可以是标题行（会被忽略）
   - 后续行的第一列将作为JSON的key，第二列作为value

2. 运行命令
```bash
pnpm start <Word文档路径> "<要查找的标题>"
```

示例：
```bash
pnpm start ./sample.docx "系统配置"
```

## 示例文档生成

项目包含一个示例文档生成器，可以创建用于测试的Word文档：

```bash
# 生成示例文档
pnpm ts-node src/create-sample-doc.ts
```

这将生成一个包含示例表格的`sample.docx`文件。

## 输出示例

```json
{
  "服务器地址": "http://localhost:8080",
  "数据库名称": "test_db",
  "用户名": "admin",
  "端口号": "5432"
}
```

## 技术栈

- TypeScript
- Node.js
- mammoth（用于读取Word文档）
- cheerio（用于解析HTML）
- docx（用于生成示例文档）

## 注意事项

- 确保Word文档中的标题文本与命令行参数完全匹配（包括空格）
- 表格必须紧跟在目标标题后面
- 表格必须是两列格式

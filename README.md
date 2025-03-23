<<<<<<< HEAD
# Excel文件合并工具

这是一个用于合并多个Excel文件的工具，专门用于处理营销现场作业计划审批表。

## 功能特点

1. 支持同时选择多个Excel文件进行合并
2. 自动验证文件结构，确保符合要求
3. 支持多表格文件处理，自动选择有效表格
4. 自动删除无效行（只有序号的行）
5. 自动标记重复数据行（浅红色背景）
6. 保持原始列宽
7. 自动调整行高以适应内容
8. 支持自定义输出模板
9. 显示处理进度
10. 自动保存到桌面
11. 文件规范检查功能
12. 智能识别表头位置（支持第4-5行表头）

## 使用说明

1. 运行程序：
   - 双击 `Excel文件合并工具.exe`
   - 程序窗口会自动居中显示

2. 选择文件：
   - 点击"选择Excel文件"按钮
   - 可以选择多个文件（最多5个）
   - 选中的文件会显示在列表中
   - 可以点击"清除选择"按钮重新选择

3. 选择输出模板：
   - 点击"选择输出模板"按钮
   - 选择要使用的模板文件
   - 合并后的数据将从第7行开始写入

4. 合并文件：
   - 点击"合并文件"按钮开始处理
   - 进度条会显示处理进度
   - 处理完成后会自动保存到桌面
   - 输出文件名格式：营销现场作业计划审批表_YYYYMMDD.xlsx

5. 规范检查：
   - 点击"规范检查"按钮打开检查窗口
   - 在检查窗口中选择要检查的Excel文件
   - 点击"规范检查"按钮开始检查
   - 检查内容包括：
     - E列为"计量用户运维一班"或"计量用户运维二班"时，检查B列和F列是否包含D列内容
     - K列为"可接受"时，检查N列是否为"否"
     - K列为"低风险"时，检查N列是否为"是"
   - 不规范内容会用黄色标记
   - 检查结果会显示在表格中

## 文件要求

1. 输入文件要求：
   - 必须是Excel文件（.xlsx或.xls格式）
   - 表头必须位于第4行或第5行
   - B列表头必须为"作业类型（内容）"
   - 必须包含以下列：
     - 序号
     - 作业类型（内容）
     - 项目管理单位/部门
     - 供电所
     - 施工单位
     - 施工地点
     - 工作开始时间
     - 工作结束时间
     - 工作负责人及电话
     - 专业
     - 基准风险等级
     - 是否需要停电
     - 施工人数
     - 是否纳入视频监督
     - 备注

2. 输出模板要求：
   - 必须是Excel文件（.xlsx格式）
   - 第6行及之前为表头
   - 从第7行开始写入数据

## 数据处理规则

1. 数据验证：
   - 自动验证文件结构
   - 智能识别表头位置（支持第4-5行表头）
   - 自动验证B列表头是否为"作业类型（内容）"
   - 对于多表格文件，自动选择第一个有效的表格
   - 如果所有表格都无效，会提示错误

2. 数据清理：
   - 自动删除只有序号的行
   - 自动删除完全空白的行
   - 自动标记重复数据行（浅红色背景）

3. 格式处理：
   - 保持原始列宽
   - 自动调整行高以适应内容
   - 时间格式统一为"YYYY/MM/DD"
   - 所有单元格居中对齐
   - 自动换行显示

4. 规范检查规则：
   - E列为"计量用户运维一班"或"计量用户运维二班"时：
     - B列必须包含D列内容
     - F列必须包含D列内容
   - K列为"可接受"时：
     - N列必须为"否"
   - K列为"低风险"时：
     - N列必须为"是"

## 注意事项

1. 程序运行需要Windows操作系统
2. 建议使用最新版本的Excel文件
3. 如果文件较大，处理可能需要一些时间
4. 请确保有足够的磁盘空间
5. 输出文件会自动保存到桌面
6. 规范检查会修改原文件，建议先备份

## 错误处理

1. 如果文件结构不正确，会显示错误提示
2. 如果处理过程中出现错误，会显示详细的错误信息
3. 如果所有表格都无效，会提示"没有找到有效的表格结构"
4. 规范检查结果会显示在检查窗口的表格中

## 技术支持

如有问题或建议，请联系技术支持。

## 依赖项

- Python 3.8+
- pandas
- openpyxl
- PySide6

## 安装依赖

```bash
pip install -r requirements.txt
```

## 更新日志

### 2024-06-30
- 优化表头识别功能，现在支持第4行或第5行的表头
- 增加B列表头名称验证，确保为"作业类型（内容）"
- 提高文件解析的兼容性和稳定性

### 2024-03-21
- 添加输出模板选择功能
- 优化序号生成逻辑，从1开始递增
- 保持模板文件原始列宽
- 添加进度条显示
- 优化时间格式显示（YYYY/MM/DD）
- 输出文件自动保存到桌面 
=======
# excel-merger-tool
文件合并工具（营销现场作业计划审批表专用）
>>>>>>> 1c0ce3d61e4c394d7739e798adef66f391d3d338

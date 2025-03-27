# Excel文件合并工具

这是一个用于合并多个Excel文件的工具，专门设计用于处理营销现场作业计划审批表。

## 功能特点

1. 支持多文件选择
   - 可以从不同目录选择文件
   - 支持多次选择，最多合并5个文件
   - 支持清除已选择的文件

2. 文件格式验证
   - 自动验证表格结构
   - 检查表头格式（第4行或第5行）
   - 验证必要列的存在性
   - 检查"作业类型（内容）"等关键列

3. 数据处理
   - 自动合并多个Excel文件
   - 按时间排序
   - 自动处理重复数据
   - 保持原始格式

4. 用户界面
   - 简洁直观的操作界面
   - 实时进度显示
   - 详细的错误提示
   - 支持文件预览和规范检查

## 使用说明

1. 选择文件
   - 点击"选择Excel文件"按钮
   - 可以多次选择不同目录的文件
   - 已选择的文件会显示在列表中
   - 使用"清除选择"按钮可以清空已选文件

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

3. 合并文件
   - 确保已选择文件和模板
   - 点击"合并文件"开始处理
   - 等待进度条完成
   - 合并后的文件将保存到桌面

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

1. 文件格式要求：
   - 必须是.xlsx或.xls格式
   - 表头必须在第4行或第5行
   - B列表头必须为"作业类型（内容）"
   - 必须包含"工作开始时间"和"工作结束时间"列

2. 使用限制：
   - 最多同时合并5个文件
   - 输出文件将自动保存到桌面
   - 输出文件名格式：附录2：营销现场作业计划审批表_当前日期.xlsx

3. 错误处理：
   - 程序会显示详细的错误信息
   - 如果部分文件合并成功，会显示警告信息
   - 可以查看详细错误信息了解具体问题

## 常见问题

1. 文件无法合并
   - 检查文件格式是否正确
   - 确保表头在正确的位置
   - 查看错误提示了解具体原因

2. 合并结果不完整
   - 检查源文件是否包含所有必要列
   - 确保时间格式正确
   - 查看是否有空行或格式错误

3. 模板相关问题
   - 确保模板文件格式正确
   - 检查模板文件是否完整
   - 必要时重新选择模板文件

## 更新日志

### 2025-03-27
- 改进文件预览功能的统计信息：
  - 修复本单位和外施工单位统计逻辑
  - 优化统计信息显示格式
  - 准确识别计量电网运行班、计量用户运维一班、计量用户运维二班为本单位
  - 其他施工单位自动归类为外施工单位

### 2025-03-25
- 添加行高限制功能，最大行高限制为180
- 优化表格显示效果
- 改进错误提示信息显示
- 添加文件预览功能，支持以下统计信息：
  - 显示文件基本信息（路径、表格数量、行数、列数）
  - 显示作业计划审批统计（日期范围、总项数、风险等级分布）
  - 显示本单位和外施工单位统计
  - 显示已发布项目数量
  - 显示各施工单位的作业数量
- 添加规范检查功能，支持以下检查：
  - 检查运维班作业内容的一致性
  - 检查风险等级与视频监督的对应关系
  - 使用黄色标记显示不规范内容
  - 在表格中显示详细的检查结果

### 2025-03-24
- 支持从不同目录选择文件
- 改进文件选择逻辑，支持多次选择文件
- 优化错误提示信息显示
- 添加文件选择清除功能
- 改进多文件处理逻辑

### 2025-03-23
- 优化表头识别功能，现在支持第4行或第5行的表头
- 增加B列表头名称验证，确保为"作业类型（内容）"
- 提高文件解析的兼容性和稳定性

### 2025-03-21
- 添加输出模板选择功能
- 优化序号生成逻辑，从1开始递增
- 保持模板文件原始列宽
- 添加进度条显示
- 优化时间格式显示（YYYY/MM/DD）
- 输出文件自动保存到桌面

# 电池测试数据处理工具

## 功能简介
本工具用于处理和分析电池测试数据，主要包括以下功能：
# Start of Selection
1. 循环数据处理
2. 中检数据处理
3. DCR（直流内阻）计算
4. 测试报告生成
5. 绘制容量和能量的散点图
6. 绘制中检DCR和DCR增长率的散点图
# End of Selection

## 主要模块

### ZPDataModule
处理中检数据相关的功能模块，包括：
- 容量和能量数据处理
- DCIR（直流内阻）计算
- DCIR Rise（内阻上升率）计算
- 数据表格生成和格式化

#### 关键功能
1. **基本数据处理**
   - 容量计算
   - 能量计算
   - 容量保持率计算
   - 能量保持率计算
   - 绘制容量和能量的散点图
   - 绘制中检DCR和DCR增长率的散点图

2. **DCIR计算**
   - 90%、50%、10% SOC点的DCIR值计算
   - 计算公式：DCIR = (搁置电压 - 放电电压) / |放电电流| * 1000

3. **DCIR Rise计算**
   - 基于首次测量值的DCIR增长率
   - 计算公式：Rise = (当前值 - 基准值) / 基准值 * 100%

### 数据处理方法
1. **容量计算方式**
   - 仅中检一次：直接使用单次测量数据
   - 三圈中检求平均值：连续三次测量取平均值

2. **DCIR测量条件**
   - 支持30s和10s放电时间点的测量
   - 自动处理不同SOC点的工步数据

## 使用说明

### 配置要求
1. **循环配置表**
   - 中检间隔圈数设置
   - 容量标定方式选择
   - SOC点工步号配置
   - 放电时间设置

2. **数据文件要求**
   - 循环数据文件（工步数据表）
   - 中检容量数据文件（工步数据表）
   - 中检DCR数据文件（详细数据表）

### 操作步骤
1. 在文件信息表中填写相关文件名
2. 设置循环配置信息
3. 点击"输出报告"按钮生成测试报告

### 输出说明
生成的报告包含：
1. 基本数据表
   - 循环圈数
   - 容量数据
   - 能量数据
   - 保持率数据

2. DCIR数据表
   - 90% SOC的DCIR值
   - 50% SOC的DCIR值
   - 10% SOC的DCIR值

3. DCIR Rise数据表
   - 各SOC点的DCIR增长率

4. 容量和能量散点图，中检DCR和DCR增长率散点图
   - 容量和能量随循环圈数的变化
   - 中检DCR和DCR增长率随循环圈数的变化

## 注意事项
1. 确保所有输入文件格式正确
2. 正确配置工步号信息
3. 数据异常时检查日志输出
4. 大数据量处理时可能需要较长时间

## 错误处理
- 所有错误信息会记录在日志中
- 主要错误类型包括：
  - 文件不存在
  - 数据格式错误
  - 工步号配置错误
  - 计算过程异常

## 维护说明
代码中包含详细的注释和错误处理，便于后续维护和更新。如需修改，请参考相应模块的注释说明。 
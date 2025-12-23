# 📦 箱贴（箱明細シール）生成功能设计方案 (V2.0)

## 1. 核心变更确认
- **数据源**：工厂实际返还的箱设定明细表格（Excel）。
- **拆分逻辑**：工厂会在表格中**显式**通过新增行或填写不同 `CTN_NO` 来体现拆分结果。
- **系统职责**：系统不再进行自动拆箱计算，而是**忠实读取**表格中的每一行（或每一组相同 `CTN_NO` 的行）作为一箱数据。

## 2. 需求规格
- **尺寸**：100mm × 120mm
- **排版**：A4 纸，每页 4 张（2行 × 2列）
- **必需字段**：
  - 店铺番号、店铺名称、店着日、部门（取管理No前3位）
  - メーカー品番、品名（暂待定，留空）、数量
  - 箱ID（模拟）、入数、箱序号（CTN_NO）

## 3. 技术实现方案

### 3.1 数据读取 (BoxSettingReader)
位于 `AutoPackage/excel_reader.py`。

- **输入**：工厂返还的 `.xlsx` 文件。
- **处理逻辑**：
  1. 遍历所有 `PT-xx` Sheet。
  2. 识别表头结构（SKU列动态解析）。
  3. 遍历数据行：
     - 读取 `CTN_NO`。
     - 读取店铺信息、日期信息。
     - 读取各 SKU 列的数量。
     - **聚合逻辑**：以 `(Store Code, CTN_NO)` 为键进行聚合。如果工厂把一箱拆成多行写（例如不同SKU分行写），需要合并；如果工厂把一行拆成多行且 `CTN_NO` 不同，则视为不同箱。
     - *当前策略*：假设每一行对应一箱，或者相同 `CTN_NO` 的行属于同一箱。为了稳健，我们将按 `CTN_NO` 聚合同一 PT Sheet 内的数据。

### 3.2 箱贴生成 (BoxLabelGenerator)
位于 `AutoPackage/box_label_generator.py`。

- **输入**：`List[BoxData]`
- **布局算法**：
  - 页面尺寸：A4 (210x297mm)
  - 标签尺寸：100x120mm
  - 边距计算：
    - X_MARGIN = (210 - 200) / 2 = 5mm
    - Y_MARGIN = (297 - 240) / 2 = 28.5mm
  - 坐标系统（ReportLab原点在左下角）：
    - (0, 1) Top-Left: x=5, y=297-28.5-120 = 148.5
    - (1, 1) Top-Right: x=105, y=148.5
    - (0, 0) Bottom-Left: x=5, y=28.5
    - (1, 0) Bottom-Right: x=105, y=28.5

### 3.3 Web 接口
位于 `web_server/web_app.py`。

- **新增接口**：`POST /api/generate-labels-from-file`
- **参数**：上传的文件对象。
- **流程**：
  1. 保存上传文件到临时目录。
  2. 调用 `BoxSettingReader` 读取数据。
  3. 调用 `BoxLabelGenerator` 生成 PDF。
  4. 返回 PDF 下载链接。

## 4. 数据结构定义

```python
class BoxData(TypedDict):
    store_code: str
    store_name: str
    ctn_no: str          # 实际箱号 (e.g., "C-001")
    pattern: str         # 原始 Pattern
    delivery_date: str   # 出区日
    store_date: str      # 店着日
    dept: str            # 部门
    kanri_no: str        # 管理No
    total_qty: int       # 总数
    items: List[BoxItem] # 明细

class BoxItem(TypedDict):
    maker_code: str      # 品番-色-尺码
    product_name: str    # 品名 (待定)
    qty: int
```

# HVAC DUCT SUPPLY AIR - AutoCAD Plugin

## 📋 Mô tả
Plugin AutoCAD để tạo và quản lý hệ thống ống dẫn khí HVAC với tick hiển thị width tự động.

## 🎯 Tính năng chính
- Vẽ polyline với tick hiển thị width
- Chỉnh sửa width và tự động cập nhật tick
- Break polyline và tự động tạo tick cho 2 polyline mới
- Load lại tick từ database

## 🚀 Cài đặt
1. Build project trong Visual Studio
2. Copy file `HVAC_DUCT.dll` vào thư mục AutoCAD
3. Load plugin trong AutoCAD bằng lệnh `NETLOAD`
4. Sử dụng các lệnh bên dưới

## 📝 Danh sách lệnh

### 1. `TAN25_HVAC_DUCT_SUPPLY_AIR` - Tạo duct
**Mục đích**: Tạo polyline mới với tick và MText hiển thị width

**Cách sử dụng**:
1. Chạy lệnh `TAN25_HVAC_DUCT_SUPPLY_AIR`
2. Nhập width (ví dụ: 12)
3. Vẽ polyline bằng cách click các điểm
4. Nhấn Enter để kết thúc
5. Tick và MText sẽ tự động hiển thị

**Kết quả**:
- Polyline có XData chứa thông tin width và length
- MText hiển thị width và length (ví dụ: "12"∅ x 24.5'")
- Tick hiển thị trên polyline

---

### 2. `TAN25_HVAC_DUCT_SUPPLY_AIR_EDIT_WIDTH` - Chỉnh sửa width
**Mục đích**: Thay đổi width của polyline có sẵn

**Cách sử dụng**:
1. Chạy lệnh `TAN25_HVAC_DUCT_SUPPLY_AIR_EDIT_WIDTH`
2. Chọn polyline cần chỉnh sửa
3. Nhập width mới
4. MText và tick sẽ tự động cập nhật

**Kết quả**:
- Width được cập nhật trong XData
- Length được tính toán lại tự động
- MText hiển thị width và length mới
- Tick được vẽ lại với width mới

---

### 3. `TAN25_HVAC_DUC_AUTOBREAK` - Break duct
**Mục đích**: Break polyline và tự động tạo tick cho 2 polyline mới

**Cách sử dụng**:
1. Chạy lệnh `TAN25_HVAC_DUC_AUTOBREAK`
2. Chọn polyline cần break
3. Chọn điểm break thứ nhất
4. Chọn điểm break thứ hai
5. Hệ thống tự động break và tạo tick

**Kết quả**:
- Polyline gốc bị xóa tick
- 2 polyline mới được tạo
- Mỗi polyline mới có XData (width, length) và MText riêng
- Length được tính toán lại cho từng polyline mới
- Tick hiển thị cho cả 2 polyline mới

---

### 4. `TAN25_HVAC_DUCT_SUPPLY_AIR_LOAD_TEMP` - Load lại tick
**Mục đích**: Tải lại tất cả tick từ database

**Cách sử dụng**:
1. Chạy lệnh `TAN25_HVAC_DUCT_SUPPLY_AIR_LOAD_TEMP`
2. Tất cả tick sẽ được load lại

**Khi nào sử dụng**:
- Sau khi mở file mới
- Khi tick không hiển thị
- Sau khi thay đổi cài đặt

## 🔧 Quy trình sử dụng

### Tạo hệ thống duct mới:
1. `TAN25_HVAC_DUCT_SUPPLY_AIR` - Tạo duct chính
2. `TAN25_HVAC_DUCT_SUPPLY_AIR` - Tạo các nhánh duct
3. `TAN25_HVAC_DUCT_SUPPLY_AIR_EDIT_WIDTH` - Điều chỉnh width nếu cần

### Chỉnh sửa hệ thống duct:
1. `TAN25_HVAC_DUCT_SUPPLY_AIR_EDIT_WIDTH` - Thay đổi width
2. `TAN25_HVAC_DUC_AUTOBREAK` - Chia nhỏ duct
3. `TAN25_HVAC_DUCT_SUPPLY_AIR_LOAD_TEMP` - Load lại tick

## 📊 Thông tin kỹ thuật

### Công nghệ API sử dụng:
- **AutoCAD .NET API**: Autodesk.AutoCAD.ApplicationServices
- **Database Services**: Autodesk.AutoCAD.DatabaseServices
- **Editor Input**: Autodesk.AutoCAD.EditorInput
- **Geometry**: Autodesk.AutoCAD.Geometry
- **Graphics Interface**: Autodesk.AutoCAD.GraphicsInterface
- **Runtime**: Autodesk.AutoCAD.Runtime

### XData Structure:
- **APP_NAME**: "HVAC_DUCT_SUPPLY_AIR"
- **Width**: Double value (width của duct)
- **Length**: Double value (độ dài của duct)
- **MText Handle**: Handle của MText liên kết

### XData Storage Details:
- **TypeCode 1001**: Application name ("HVAC_DUCT_SUPPLY_AIR")
- **TypeCode 1040**: Width value (Double)
- **TypeCode 1040**: Length value (Double) 
- **TypeCode 1005**: MText Handle (Long)

### Layer:
- **M-ANNO-TAG-DUCT**: Layer cho MText (màu 50)

### Tick Display:
- Tick được vẽ bằng FilletOverrule
- Màu đỏ cho đoạn ngắn
- Màu xanh cho đoạn dài (theo width)
- Khoảng cách tick: 4.0 units

### API Components:
- **DrawableOverrule**: Vẽ tick tùy chỉnh trên polyline
- **Transaction Management**: Quản lý database operations
- **XData System**: Lưu trữ dữ liệu tùy chỉnh (width, length, MText)
- **MText API**: Tạo và quản lý text annotations
- **Entity Selection**: Chọn và thao tác với entities
- **Polyline Length Calculation**: Tính toán độ dài polyline tự động

### Length Calculation:
- **Method**: `Polyline.GetDistanceAtParameter()`
- **Storage**: Lưu vào XData với TypeCode 1040
- **Unit**: Theo đơn vị của drawing (inch/mm)
- **Update**: Tự động cập nhật khi polyline thay đổi

## 🐛 Xử lý sự cố

### Tick không hiển thị:
1. Chạy `TAN25_HVAC_DUCT_SUPPLY_AIR_LOAD_TEMP`
2. Kiểm tra polyline có XData không
3. Restart AutoCAD nếu cần

### MText không cập nhật:
1. Kiểm tra polyline có XData đúng không
2. Chạy lại lệnh edit width
3. Kiểm tra layer M-ANNO-TAG-DUCT

### Break không hoạt động:
1. Đảm bảo polyline có XData
2. Chọn đúng 2 điểm break
3. Chạy `TAN25_HVAC_DUCT_SUPPLY_AIR_LOAD_TEMP` sau break

## 📁 File structure
```
HVAC_DUCT/
├── HVAC_DUCT/
│   ├── Supply_duct.cs          # Main code
│   ├── HVAC_DUCT.csproj        # Project file
│   └── packages.config         # Dependencies
├── packages/                   # NuGet packages
└── README.md                   # This file
```

## 🔄 Version History
- **v1.0**: Tạo duct với tick cơ bản
- **v2.0**: Thêm chức năng edit width
- **v3.0**: Thêm chức năng break duct
- **v4.0**: Tối ưu hóa code, xóa function không cần thiết

## 👨‍💻 Developer

### Thông tin dự án:
- **Author**: TAN2025
- **Namespace**: TAN2025_HVAC_DUCT_SUPPLY_AIR
- **Target Framework**: .NET Framework 4.7
- **AutoCAD Version**: 2024

### Dependencies:
- **AutoCAD.NET.24.1.51000**: Core AutoCAD .NET API
- **AutoCAD.NET.Core.24.1.51000**: Core services
- **AutoCAD.NET.Model.24.1.51000**: Database model services

### API Usage Details:

#### **ApplicationServices**:
- `Application.DocumentManager.MdiActiveDocument`: Quản lý document
- `Document.Editor`: Thao tác với command line
- `Database`: Truy cập database AutoCAD

#### **DatabaseServices**:
- `Transaction`: Quản lý database operations
- `BlockTable/BlockTableRecord`: Quản lý entities
- `Polyline`: Entity chính được sử dụng
- `MText`: Text annotations
- `RegAppTable`: Đăng ký ứng dụng cho XData

#### **EditorInput**:
- `PromptEntityOptions/Result`: Chọn entities
- `PromptPointOptions/Result`: Chọn điểm
- `PromptDoubleOptions/Result`: Nhập số

#### **Geometry**:
- `Point3d`: Tọa độ 3D
- `Vector3d`: Vector tính toán
- `Matrix3d`: Transformations

#### **GraphicsInterface**:
- `DrawableOverrule`: Vẽ custom graphics
- `DrawableOverrule.Draw()`: Vẽ tick tùy chỉnh

#### **Runtime**:
- `CommandMethod`: Định nghĩa AutoCAD commands
- `IExtensionApplication`: Lifecycle management

## 📞 Hỗ trợ
Nếu gặp vấn đề, vui lòng kiểm tra:
1. AutoCAD version tương thích
2. .NET Framework đã cài đặt
3. Plugin đã được load đúng cách
4. File không bị corrupt

---
**Lưu ý**: Plugin này được thiết kế cho AutoCAD 2024 và .NET Framework 4.7. Đảm bảo môi trường phù hợp trước khi sử dụng.
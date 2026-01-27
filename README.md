<div align="center">


# 🛠️ SuperTools


### Add-in chuyển đổi số thành chữ và tiện ích cho Excel

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.7-purple.svg)](https://dotnet.microsoft.com/)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.9.0-green.svg)](https://excel-dna.net/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

**Một dự án phi lợi nhuận nhằm hỗ trợ cộng đồng người dùng Excel Việt Nam**

[Tính năng](#-tính-năng) • [Cài đặt](#-cài-đặt) • [Sử dụng](#-sử-dụng) • [Đóng góp](#-đóng-góp) • [Giấy phép](#-giấy-phép)

</div>

---

## 📖 Giới thiệu

**SuperTools** là một Excel Add-in mã nguồn mở, miễn phí, cung cấp các hàm chuyển đổi số thành chữ tiếng Việt, tiếng Anh và các tiện ích xử lý chuỗi cho Excel.

- ✅ **Miễn phí** - Hoàn toàn không thu phí
- ✅ **Mã nguồn mở** - Minh bạch, có thể cải tiến
- ✅ **Phi lợi nhuận** - Phục vụ cộng đồng
- ✅ **Dễ sử dụng** - Tích hợp trực tiếp vào Excel
- ✅ **Chuẩn hóa** - Tuân thủ quy tắc tiếng Việt

## ✨ Tính năng


### 🇻🇳 Chuyển đổi số thành chữ tiếng Việt

| Hàm                | Mô tả                                 | Ví dụ                                                                                         |
| ------------------ | ------------------------------------- | --------------------------------------------------------------------------------------------- |
| `=VND(số)`         | Chuyển số thành chữ tiền tệ Việt Nam  | `=VND(1234567)` → "Một triệu hai trăm ba mươi bốn nghìn năm trăm sáu mươi bảy đồng chẵn."    |
| `=REMOVEACCENT(chuỗi)` | Loại bỏ dấu tiếng Việt khỏi chuỗi | `=REMOVEACCENT("Trần Thị Bích Ngọc")` → "Tran Thi Bich Ngoc"                                |


### 🇺🇸 Chuyển đổi số thành chữ tiếng Anh

| Hàm         | Mô tả                        | Ví dụ                                                                                       |
| ----------- | ---------------------------- | ------------------------------------------------------------------------------------------- |
| `=USD(số)`  | Chuyển số thành chữ tiền USD | `=USD(1234.56)` → "One thousand two hundred thirty-four dollars and fifty-six cents."       |


### 🔧 Tiện ích bổ sung

| Hàm                | Mô tả                        | Ví dụ                                             |
| ------------------ | ---------------------------- | ------------------------------------------------- |
| `=REMOVEACCENT(chuỗi)` | Loại bỏ dấu tiếng Việt      | `=REMOVEACCENT("Nguyễn Văn An")` → "Nguyen Van An" |

### 🎯 Tùy chọn linh hoạt


Hàm `VND` và `USD` hỗ trợ các tham số tùy chọn:

```excel
=VND(số, [có_đơn_vị], [đơn_vị_nghìn])
=USD(số, [có_đơn_vị])
```

**Ví dụ:**

- `=VND(1500000)` → "Một triệu năm trăm nghìn đồng chẵn."
- `=VND(1500000, TRUE, FALSE)` → "Một triệu năm trăm ngàn đồng chẵn."
- `=USD(1234.56)` → "One thousand two hundred thirty-four dollars and fifty-six cents."

## 🚀 Cài đặt

### Yêu cầu hệ thống

- Windows 7 trở lên
- Microsoft Excel 2013 trở lên
- .NET Framework 4.7 hoặc cao hơn

### Hướng dẫn cài đặt

1. **Tải về phiên bản mới nhất**

   - Truy cập [Releases](../../releases) và tải file `.xll`.

2. **Cài đặt Add-in**

   - Mở file `.xll` vừa tải về (thường nằm trong phần `Downloads`)
   - Click chuột phải vào file vừa tải về, chọn `Properties`.
   - Tick vào ô `Unlock` trong thẻ *General* rồi nhấn `OK` (Nếu có)
   - Mở Excel
   - Vào **File** → **Options** → **Add-ins**
   - Chọn **Excel Add-ins** và nhấn **Go...**
   - Nhấn **Browse...** và chọn file đã tải
   - Tick vào **SuperTools Add-In** và nhấn **OK**

3. **Kiểm tra**
   - Mở Excel và thử hàm `=VND(12345)`
   - Nếu hiển thị "Mười hai nghìn ba trăm bốn mươi lăm đồng chẵn" → Thành công! 🎉

## 📚 Sử dụng

### Ví dụ cơ bản


#### Chuyển đổi số thành chữ tiền Việt

```excel
A1: 1234567
B1: =VND(A1,0)
→ Kết quả: "Một triệu hai trăm ba mươi bốn nghìn năm trăm sáu mươi bảy."
```


#### Chuyển đổi số thập phân

```excel
A1: 1234.56
B1: =VND(A1)
→ Kết quả: "Một nghìn hai trăm ba mươi bốn đồng năm mươi sáu xu."
```


#### Chuyển đổi số âm

```excel
A1: -500000
B1: =VND(A1,1,0)
→ Kết quả: "Âm năm trăm ngàn đồng chẵn."
```


#### Chuyển đổi sang tiếng Anh

```excel
A1: 1234.56
B1: =USD(A1)
→ Kết quả: "One thousand two hundred thirty-four dollars and fifty-six cents."
```

### Ví dụ nâng cao

#### Sử dụng trong hóa đơn

```excel
A1: 15750000
B1: =VND(A1, TRUE, TRUE)
→ "Mười lăm triệu bảy trăm năm mươi nghìn đồng chẵn"
```

## 🛠️ Phát triển


### Công nghệ sử dụng

- **Ngôn ngữ**: C# (.NET Framework 4.8)
- **Add-in Engine**: [Excel-DNA](https://excel-dna.net/) 1.9.0
- **IDE**: Visual Studio 2019/2022
- **Hệ điều hành**: Windows

### Build từ mã nguồn

```bash
# Clone repository
git clone https://github.com/your-username/SuperTools.git
cd SuperTools

# Mở solution
SuperTools.slnx

# Build trong Visual Studio (Ctrl+Shift+B)
# Output: SuperTools\bin\Debug\SuperTools-AddIn.xll
```

### Cấu trúc dự án

```
SuperTools/
├── SuperTools/
│   ├── Functions.cs           # Các hàm Excel chính
│   ├── Helper.cs              # Hàm phụ trợ
│   ├── SuperTools.csproj      # Project configuration
│   └── SuperTools-AddIn.dna   # Excel-DNA manifest
├── packages/                  # NuGet packages
├── README.md                  # Tài liệu này
└── LICENSE                    # Giấy phép MIT
```

### Đóng góp mã nguồn

Chúng tôi rất hoan nghênh mọi đóng góp! Để đóng góp:

1. **Fork** repository này
2. Tạo **branch** mới (`git checkout -b feature/TinhNangMoi`)
3. **Commit** thay đổi (`git commit -m 'Thêm tính năng mới'`)
4. **Push** lên branch (`git push origin feature/TinhNangMoi`)
5. Tạo **Pull Request**

### Coding Guidelines

- Tuân thủ C# coding conventions
- Comment rõ ràng cho các hàm phức tạp
- Viết unit tests cho các tính năng mới
- Đảm bảo backward compatibility

## 🤝 Đóng góp

### Báo lỗi

Nếu bạn phát hiện lỗi, vui lòng [tạo issue](../../issues/new) với thông tin:

- **Mô tả lỗi**: Lỗi xảy ra như thế nào?
- **Các bước tái hiện**: Làm thế nào để gặp lỗi?
- **Môi trường**: Windows version, Excel version
- **Screenshot**: Nếu có thể

### Đề xuất tính năng

Có ý tưởng mới? [Tạo feature request](../../issues/new) với:

- **Mô tả tính năng**: Tính năng làm gì?
- **Use case**: Sử dụng trong trường hợp nào?
- **Ví dụ**: Cách sử dụng mong muốn

### Hỗ trợ tài chính

Dự án này hoàn toàn phi lợi nhuận và miễn phí. Nếu bạn thấy hữu ích, bạn có thể:

- ⭐ **Star** repository này
- 📢 **Chia sẻ** với đồng nghiệp
- 💡 **Đóng góp** mã nguồn hoặc ý tưởng

## 📄 Giấy phép

Dự án này được phát hành dưới giấy phép **MIT License** - xem file [LICENSE](LICENSE) để biết chi tiết.

```
MIT License

Copyright (c) 2024 NienTools Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 🙏 Lời cảm ơn

Dự án này được lấy cảm hứng từ:

- **vmtools** - Công cụ tiên phong trong lĩnh vực chuyển đổi số thành chữ tiếng Việt
- **Excel-DNA** - Framework tuyệt vời cho Excel Add-in development
- **Cộng đồng Excel Việt Nam** - Động lực phát triển dự án

## 📞 Liên hệ

- **Email**: [niennguyen@nien.edu.vn](mailto:niennguyen@nien.edu.vn)

---

<div align="center">

**Được phát triển với ❤️ bởi cộng đồng Việt Nam**

[⬆ Về đầu trang](#-nientools)

</div>

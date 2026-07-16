# Mạch Tiền

Ứng dụng web xử lý sao kê ngân hàng ngay trong trình duyệt:

- đọc nhiều file Excel hoặc CSV;
- tự nhận diện và cho phép chỉnh cách khớp cột;
- hợp nhất giao dịch thành một bảng chuẩn;
- tóm tắt thu, chi, dòng tiền ròng và nhóm chi phí;
- xuất bảng đầy đủ, mẫu gọn hoặc điền vào một file Excel có sẵn.

Dữ liệu file chỉ được xử lý trên thiết bị của người dùng.

## Yêu cầu

- Node.js `>=22.13.0`

## Chạy tại máy

```bash
npm install
npm run dev
```

Mã giao diện và logic xử lý nằm trong `app/page.tsx`; hệ thống hình ảnh nằm trong `app/globals.css`.

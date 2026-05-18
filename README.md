# Jctrans Data Scrape Bot

Bot Selenium tự động scrape danh bạ công ty trên
[`https://www.jctrans.com/en/directory/`](https://www.jctrans.com/en/directory/)
và xuất ra Excel.

## Tính năng

- Tự đăng nhập jctrans.com (xử lý popup "account is logged in elsewhere").
- Duyệt toàn bộ dropdown 236 quốc gia (hoặc 1 quốc gia chỉ định bằng `--country`).
- Với mỗi quốc gia: bấm Search → đi qua từng trang → click vào từng công ty.
- Trên trang chi tiết, trích:
  - Country, Company name, Member ID, Year
  - Location, Website
  - Main Business / Sea Freight Advantageous / Air Freight Advantageous (chip nối bởi `|`)
  - Contact list: với mỗi contact lấy Name, Position, Email, Phone, WeChat, WhatsApp, Skype.
    Loại field được nhận diện qua icon (`<g id="icon/mail">`, `icon/phone`, …) chứ
    không phụ thuộc vị trí, nên đúng ngay cả khi công ty thiếu một số kênh liên hệ.
- **Bỏ qua** các công ty hiển thị `It is NOT a JCtrans member` hoặc `MEMBERSHIP SUSPEND`.
- Mỗi contact tách thành 1 row riêng (cột A–J lặp lại, cột K–Q là dữ liệu contact).
- Có checkpoint (`jctrans_checkpoint.json`) ghi các country đã hoàn tất → khi chạy
  lại với `--resume` sẽ bỏ qua những country đã làm.
- Excel auto-save sau mỗi công ty.

## Cài đặt

Yêu cầu Python 3.10+ và Google Chrome (hoặc Microsoft Edge build mới — selenium-manager
tự tải đúng driver).

```bash
git clone https://github.com/BrentLam123/Jctrans-data-scrapebot.git
cd Jctrans-data-scrapebot
python -m venv .venv && source .venv/bin/activate     # Linux/Mac
# .\.venv\Scripts\activate                            # Windows PowerShell
pip install -r requirements.txt
cp .env.example .env  # rồi sửa JCT_EMAIL / JCT_PASSWORD nếu cần
```

## Cấu hình

Thông tin đăng nhập đọc qua env vars (hoặc giữ giá trị mặc định trong `scraper.py`).

| Biến môi trường   | Mặc định                 | Ý nghĩa                              |
|-------------------|--------------------------|--------------------------------------|
| `JCT_EMAIL`       | `brent@pio-logistics.vn` | Email/Username đăng nhập             |
| `JCT_PASSWORD`    | `0925587314aA`           | Mật khẩu                             |
| `JCT_EXCEL`       | `jctrans_data.xlsx`      | Đường dẫn file Excel xuất ra         |
| `JCT_CHECKPOINT`  | `jctrans_checkpoint.json`| Đường dẫn file checkpoint            |
| `CHROME_BIN`      | *(auto)*                 | Override đường dẫn Chrome nếu khác   |

Có thể `export JCT_EMAIL=… JCT_PASSWORD=…` trước khi chạy, hoặc dùng `.env` cùng `python-dotenv`.

## Cách chạy

Scrape toàn bộ (mặc định, headless):

```bash
python scraper.py
```

Chỉ chạy 1 quốc gia (tốt cho test):

```bash
python scraper.py --country Japan
```

Chạy nhiều quốc gia:

```bash
python scraper.py --country Japan India "United States"
```

Tiếp tục từ lần chạy trước (bỏ qua những country đã có trong checkpoint):

```bash
python scraper.py --resume
```

Chế độ test nhanh (vd. 2 công ty cho mỗi quốc gia):

```bash
python scraper.py --country Japan --max-companies 2
```

Hiển thị browser thay vì headless (chỉ chạy được trên máy local có GUI, không
phải trên VM headless):

```bash
python scraper.py --country Japan --no-headless
```

## Cấu trúc cột Excel

| Cột | Tên              | Ghi chú                                                |
|-----|------------------|--------------------------------------------------------|
| A   | STT              | Số thứ tự chạy                                         |
| B   | Country          |                                                        |
| C   | Company name     |                                                        |
| D   | Member ID        | Trống nếu không phải JCtrans member                   |
| E   | Year             | Ví dụ `15-Year`                                        |
| F   | Location         |                                                        |
| G   | Website          |                                                        |
| H   | Main Business    | Chip nối bởi ` | `                                     |
| I   | Sea Freight Adv. | Chip nối bởi ` | `                                     |
| J   | Air Freight Adv. | Chip nối bởi ` | `                                     |
| K   | Contact name     | 1 row / contact (A–J lặp lại)                         |
| L   | Position         |                                                        |
| M   | Email            | Chỉ ghi nếu chip có icon `mail`                       |
| N   | Phone            | icon `phone`                                          |
| O   | WeChat           | icon `wechat`                                         |
| P   | WhatsApp         | icon `whatsapp`                                       |
| Q   | Skype            | icon `skype`                                          |

## Logs / Debug

- Log từng phiên chạy lưu ở `logs/scrape_YYYYMMDD_HHMMSS.log`.
- Excel được lưu lại sau mỗi công ty (tránh mất dữ liệu khi crash).
- Checkpoint `jctrans_checkpoint.json` cập nhật sau mỗi country.

## Giới hạn đã biết

- Nếu tài khoản hết quyền xem contact (free / suspended), email/phone trên trang
  jctrans hiển thị mask (`pr********`). Bot ghi đúng những gì hiện ra trên trang
  (giống user thấy).
- Nếu mạng chậm hoặc trang lazy-load chưa kịp, có thể tăng `time.sleep` trong
  `parse_company` (mặc định cuộn trang 6 lần).
- Captcha: jctrans hiện không yêu cầu captcha cho login; nếu sau này bật captcha,
  cần đăng nhập tay 1 lần và tái sử dụng cookie (chưa hỗ trợ).

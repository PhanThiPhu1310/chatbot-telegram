from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from datetime import datetime
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ssl
import os
import re
import logging
import pytz

# Cấu hình logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Token từ BotFather
TOKEN = '7730025615:AAH7DHk9FQkiIo2nxheNmqGhVTgNhCK-D_w'

# Cấu hình email
EMAIL_SENDER = "phuhktmmen@gmail.com"
EMAIL_PASSWORD = "wogl ivad bznd exhx"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# Đọc dữ liệu từ file promotions.xlsx
try:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    promo_file_path = os.path.join(current_dir, 'promotions.xlsx')
    print(f"Đang tìm file khuyến mãi tại: {promo_file_path}")

    if not os.path.exists(promo_file_path):
        print(f"File promotions.xlsx không tồn tại. Tạo file mặc định với không có khuyến mãi.")
        df_promo = pd.DataFrame(
            columns=['Mã khuyến mãi', 'Loại khuyến mãi', 'Giá trị', 'Áp dụng cho', 'Ngày bắt đầu', 'Ngày kết thúc',
                     'Điều kiện áp dụng'])
    else:
        # Đọc file Excel, ép cột Ngày bắt đầu và Ngày kết thúc thành chuỗi
        df_promo = pd.read_excel(promo_file_path, dtype={'Ngày bắt đầu': str, 'Ngày kết thúc': str})

        # In dữ liệu trước khi chuyển đổi để debug
        print("Dữ liệu trước khi chuyển đổi (toàn bộ):\n", df_promo.to_string())

        # Hàm chuyển đổi ngày tháng với nhiều định dạng
        def parse_date(date_str):
            if pd.isna(date_str) or date_str == '':
                return pd.NaT
            try:
                return pd.to_datetime(date_str, format='%d/%m/%Y', errors='coerce')
            except:
                try:
                    return pd.to_datetime(date_str, format='%Y-%m-%d', errors='coerce')
                except:
                    return pd.to_datetime(date_str, errors='coerce')

        # Chuyển đổi ngày tháng
        df_promo['Ngày bắt đầu'] = df_promo['Ngày bắt đầu'].apply(parse_date)
        df_promo['Ngày kết thúc'] = df_promo['Ngày kết thúc'].apply(parse_date)

        # In toàn bộ dữ liệu sau khi chuyển đổi để kiểm tra
        print("Dữ liệu sau khi chuyển đổi (toàn bộ):\n", df_promo.to_string())
        print("Kiểu dữ liệu:\n", df_promo.dtypes)

        # Kiểm tra dữ liệu không hợp lệ
        invalid_dates = df_promo[df_promo['Ngày bắt đầu'].isna() | df_promo['Ngày kết thúc'].isna()]
        if not invalid_dates.empty:
            print("Cảnh báo: Các dòng sau trong file promotions.xlsx có ngày không hợp lệ:")
            print(invalid_dates[['Mã khuyến mãi', 'Ngày bắt đầu', 'Ngày kết thúc']])
    print("Đã đọc file khuyến mãi thành công!")
except Exception as e:
    print(f"Lỗi khi đọc file khuyến mãi: {e}")
    df_promo = pd.DataFrame(
        columns=['Mã khuyến mãi', 'Loại khuyến mãi', 'Giá trị', 'Áp dụng cho', 'Ngày bắt đầu', 'Ngày kết thúc',
                 'Điều kiện áp dụng'])

# Đọc dữ liệu từ file Excel product
try:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, 'products.xlsx')
    print(f"Đang tìm file tại: {file_path}")

    if not os.path.exists(file_path):
        print(f"File không tồn tại tại {file_path}. Vui lòng kiểm tra lại tên file!")
        exit(1)

    df = pd.read_excel(file_path)
    print("Đã đọc file thành công!")
except FileNotFoundError:
    print(f"Không tìm thấy file products.xlsx tại {file_path}. Vui lòng kiểm tra lại đường dẫn hoặc file!")
    exit(1)
except Exception as e:
    print(f"Lỗi khi đọc file: {e}")
    exit(1)

# Hàm chuẩn hóa chuỗi để so sánh (bỏ khoảng cách, dấu "-", và chuyển thành chữ thường)
def normalize_string(s):
    return s.lower().replace(" ", "").replace("-", "")

# Tạo từ điển ánh xạ từ dạng chuẩn hóa sang tên gốc
category_mapping = {normalize_string(cat): cat for cat in df['Tên sản phẩm'].unique()}

# Hàm gửi email
async def send_email(to_email, subject, body):
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = EMAIL_SENDER
        msg['To'] = to_email
        part1 = MIMEText(body, 'html')
        msg.attach(part1)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)
        logger.info(f"Email sent successfully to {to_email}")
        return True
    except Exception as e:
        logger.error(f"Failed to send email to {to_email}: {e}")
        return False

async def log_user_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message and update.message.text:
        tz = pytz.timezone('Asia/Ho_Chi_Minh')
        now = datetime.now(tz)
        if 'conversation' not in context.user_data:
            context.user_data['conversation'] = []
        context.user_data['conversation'].append(('user', update.message.text, now))
        print(f"Logged user message: {update.message.text} at {now}")

async def send_and_log(update: Update, context: ContextTypes.DEFAULT_TYPE, text: str):
    await update.message.reply_text(text)
    tz = pytz.timezone('Asia/Ho_Chi_Minh')
    now = datetime.now(tz)
    if 'conversation' not in context.user_data:
        context.user_data['conversation'] = []
    context.user_data['conversation'].append(('bot', text, now))
    print(f"Logged bot message: {text} at {now}")

# Hàm xử lý lệnh /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_msg = (
        "Chào mừng bạn đến với ChatBot của KLC! Tôi có thể hỗ trợ bạn về các sản phẩm mà công ty đang cung cấp.\n"
        "Hoặc gõ /help để xem thêm gợi ý các câu hỏi!"
    )
    await send_and_log(update, context, welcome_msg)

# Hàm xử lý lệnh /help
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_msg = (
        "CÁC CÂU HỎI THƯỜNG GẶP\n"
        "- Công ty đang cung cấp các loại sản phẩm nào?\n"
        "- Các loại Video Wall mà công ty đang cung cấp?\n"
        "- KVM Switch nào có giá dưới 10 triệu?\n"
        "- Pro-AV nào phù hợp cho hội nghị?\n"
        "- Có chương trình khuyến mãi nào cho Video Wall không?\n"
        "- Có ưu đãi nào cho phòng họp lớn không?\n"
        "Ngoài ra, tôi cũng hỗ trợ các lệnh:\n"
        "- /start: Bắt đầu lại\n"
        "- /products: Xem danh sách danh mục sản phẩm\n"
        "- /sendemail: Gửi thông tin qua email\n"
    )

    await send_and_log(update, context, help_msg)
    context.user_data['last_response'] = help_msg

# Hàm xử lý lệnh /sendemail
async def sendemail(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['email_request'] = True
    await send_and_log(update, context,
                       "Bạn muốn gửi thông tin gì qua email?\n1. Thông tin sản phẩm\n2. Chương trình khuyến mãi\n"
                       "3. Lịch sử trò chuyện\n4. Tất cả\nVui lòng chọn số từ 1 đến 4.")

# Hàm xử lý lệnh /products
async def products(update: Update, context: ContextTypes.DEFAULT_TYPE):
    categories = df['Tên sản phẩm'].unique().tolist()
    response = "Các danh mục sản phẩm của KLC:\n" + "\n".join([f"- {cat}" for cat in categories])
    await send_and_log(update, context, response)

# Hàm tìm kiếm sản phẩm theo danh mục và trả về danh sách mã sản phẩm
def get_product_codes_by_category(category):
    normalized_category = normalize_string(category)
    original_category = category_mapping.get(normalized_category)

    if not original_category:
        return None

    matched_products = df[df['Tên sản phẩm'] == original_category]
    if not matched_products.empty:
        codes = matched_products['Mã sản phẩm'].tolist()
        return f"{original_category}: {', '.join(codes)}"
    return None

# Hàm tìm kiếm sản phẩm theo giá
def search_by_price(category, price_limit, below=True):
    normalized_category = normalize_string(category)
    original_category = category_mapping.get(normalized_category)

    if not original_category:
        return f"Không tìm thấy danh mục {category}."

    matched_products = df[df['Tên sản phẩm'] == original_category]
    if below:
        matched_products = matched_products[matched_products['Giá (VND)'] < price_limit]
    else:
        matched_products = matched_products[matched_products['Giá (VND)'] > price_limit]

    if not matched_products.empty:
        response = f"Các sản phẩm {original_category} {'dưới' if below else 'trên'} {price_limit:,} VND:\n\n"
        for _, row in matched_products.iterrows():
            response += f"- {row['Mã sản phẩm']}: {row['Giá (VND)']:,} VND\n"
        return response
    return f"Xin lỗi bạn. Hiện tại không có sản phẩm {original_category} nào {'dưới' if below else 'trên'} {price_limit:,} VND."

# Hàm tìm kiếm sản phẩm theo mục đích (VD: hội nghị)
def search_by_purpose(category, purpose):
    normalized_category = normalize_string(category)
    original_category = category_mapping.get(normalized_category)

    if not original_category:
        return f"Không tìm thấy danh mục {category}."

    matched_products = df[
        (df['Tên sản phẩm'] == original_category) &
        (df['Mô tả'].str.lower().str.contains(purpose, na=False))
        ]

    if not matched_products.empty:
        response = f"Thiết bị {original_category} phù hợp cho {purpose} đây!\n\n"

        # Danh sách để lưu trữ thông tin sản phẩm
        product_highlights = []

        for _, row in matched_products.iterrows():
            product_code = row['Mã sản phẩm']
            description = row['Mô tả']

            # Tóm tắt và diễn đạt lại mô tả
            highlight = generate_highlight(description, purpose, product_code, style="practical_friendly")
            product_highlights.append((product_code, highlight))

        # Thêm thông tin sản phẩm vào response
        for product_code, highlight in product_highlights:
            response += (
                f"**{product_code}**  \n"
                f"- **Điểm nổi bật**: {highlight}  \n\n"
            )

        # Thêm phần gợi ý
        if len(product_highlights) > 1:
            response += "**Gợi ý chọn thiết bị**:  \n"
            if purpose.lower() == "hội nghị":
                for product_code, highlight in product_highlights:
                    desc_lower = highlight.lower()
                    if "phòng họp lớn" in desc_lower or "sự kiện trực tiếp" in desc_lower:
                        response += f"- Dùng cho hội nghị lớn: {product_code} là lựa chọn tốt nhờ khả năng hỗ trợ hội nghị quy mô lớn.  \n"
                    if "đa năng" in desc_lower or "dễ sử dụng" in desc_lower:
                        response += f"- Dùng cho hội nghị vừa và nhỏ: {product_code} sẽ đáp ứng tốt với tính năng đa năng và dễ sử dụng.  \n"
            else:
                response += (
                    f"- Dùng cho {purpose} quy mô lớn: Chọn sản phẩm có hiệu suất cao và hỗ trợ nhiều kết nối.  \n"
                    f"- Dùng cho {purpose} vừa và nhỏ: Chọn sản phẩm dễ sử dụng và linh hoạt.  \n"
                )
            response += f"- **Lưu ý**: Hãy xem bạn cần {purpose} ở quy mô nào để chọn thiết bị phù hợp nhé!  "
        else:
            # Khi chỉ có 1 sản phẩm
            product_code, highlight = product_highlights[0]
            response += f"**Gợi ý sử dụng**:  \n"
            desc_lower = highlight.lower()
            if "phòng họp lớn" in desc_lower or "sự kiện trực tiếp" in desc_lower:
                response += f"- {product_code} rất phù hợp cho {purpose} quy mô lớn, như phòng họp lớn hoặc sự kiện trực tiếp.  \n"
            elif "đa năng" in desc_lower or "dễ sử dụng" in desc_lower:
                response += f"- {product_code} phù hợp cho {purpose} vừa và nhỏ, nhờ tính năng đa năng và dễ sử dụng.  \n"
            else:
                response += f"- {product_code} là lựa chọn tốt cho {purpose}, bạn có thể dùng trong nhiều tình huống khác nhau.  \n"
            response += f"**Lưu ý**: Hãy xem bạn cần {purpose} ở quy mô nào để tận dụng thiết bị này tốt nhất nhé!  "

        return response

    return f"Xin lỗi bạn. Hiện tại không có sản phẩm {original_category} phù hợp cho {purpose}."


# Hàm phụ để tóm tắt và diễn đạt lại mô tả
def generate_highlight(description, purpose, product_code, style="default"):
    # Chuyển mô tả thành chữ thường để dễ xử lý
    desc_lower = description.lower()

    # Tìm các từ khóa quan trọng liên quan đến mục đích
    purpose_related = f"Phù hợp cho {purpose}" in description or purpose in desc_lower

    # Tìm các tính năng nổi bật trong mô tả
    features = []
    if "độ trễ thấp" in desc_lower or "độ trễ cực thấp" in desc_lower:
        features.append("độ trễ thấp, đảm bảo hiệu suất mượt mà")
    if "hdmi 2.0" in desc_lower:
        features.append("hỗ trợ HDMI 2.0, kết nối tiện lợi")
    if "displayport" in desc_lower:
        features.append("hỗ trợ DisplayPort, kết nối tiện lợi")
    if "đa năng" in desc_lower or "nhiều định dạng" in desc_lower:
        features.append("dùng được cho nhiều mục đích")
    if "tự động" in desc_lower:
        features.append("có tính năng tự động, dễ sử dụng")
    if "cao cấp" in desc_lower:
        features.append("cao cấp")
    if "phòng họp lớn" in desc_lower:
        features.append("phù hợp cho phòng họp lớn")
    if "sự kiện trực tiếp" in desc_lower:
        features.append("phù hợp cho sự kiện lớn")
    if "4k" in desc_lower:
        features.append("hỗ trợ 4K")
    if "chuyển đổi nhanh" in desc_lower or "phím nóng" in desc_lower:
        features.append("có thể chuyển đổi nhanh qua phím nóng")
    if "trung tâm điều hành" in desc_lower:
        features.append("rất phù hợp cho trung tâm điều hành")
    if "livestream" in desc_lower:
        features.append("phù hợp cho livestream")
    if "quảng cáo" in desc_lower:
        features.append("phù hợp cho quảng cáo")
    if "hình ảnh sắc nét" in desc_lower:
        features.append("hình ảnh sắc nét")

    # Tạo câu "điểm nổi bật" dựa trên phong cách
    if style == "practical_friendly":
        if features:
            # Kết hợp các tính năng thành câu ngắn gọn
            highlight = f"{' '.join(features).capitalize()}"
            if "bộ kvm" in desc_lower:
                highlight = f"Bộ KVM {highlight.lower()}"
            elif "bộ chia tín hiệu" in desc_lower:
                highlight = f"Bộ chia tín hiệu {highlight.lower()}"
        else:
            # Nếu không tìm thấy tính năng cụ thể, tóm tắt mô tả
            highlight = description[:150] + "..." if len(description) > 150 else description

        # Thêm mục đích sử dụng nếu có
        if "trung tâm điều hành" in desc_lower:
            highlight += ", rất phù hợp cho trung tâm điều hành"
        elif "phòng họp lớn" in desc_lower:
            highlight += ", rất phù hợp cho phòng họp lớn"
        elif "livestream" in desc_lower:
            highlight += ", phù hợp cho livestream"
        elif "quảng cáo" in desc_lower:
            highlight += ", phù hợp cho quảng cáo"
        elif "sự kiện lớn" in desc_lower:
            highlight += ", phù hợp cho sự kiện lớn"

        # Kết thúc câu
        if purpose_related:
            highlight += f" cần chất lượng hình ảnh cao."
        else:
            highlight += f"."

    elif style == "friendly_new":
        if features:
            highlight = f"{product_code} là một thiết bị rất đáng chú ý đó! Nó {', '.join(features)}."
        else:
            highlight = description[:150] + "..." if len(description) > 150 else description
            highlight = f"{product_code} có {highlight} nè."

        if purpose_related:
            highlight += f" Dùng cho các nhu cầu {purpose} chuyên nghiệp thì thật sự là một lựa chọn tuyệt vời!"
        else:
            highlight += f" Một lựa chọn rất đáng để thử cho nhu cầu {purpose} của bạn đó!"

    elif style == "friendly":
        if features:
            highlight = f"Bạn ơi, {product_code} nổi bật với {', '.join(features)}."
        else:
            highlight = description[:100] + "..." if len(description) > 100 else description
            highlight = f"Bạn ơi, {product_code} có {highlight} nè."

        if purpose_related:
            highlight += f" Dùng cho {purpose} chuyên nghiệp thì hết ý luôn!"
        else:
            highlight += f" Một lựa chọn đáng thử cho {purpose} của bạn đó!"

    else:
        # Phong cách mặc định
        if features:
            highlight = f"Thiết bị {product_code} nổi bật với {', '.join(features)}."
        else:
            highlight = description[:100] + "..." if len(description) > 100 else description

        if purpose_related:
            highlight += f" Đặc biệt phù hợp cho {purpose} chuyên nghiệp!"
        else:
            highlight += f" Một lựa chọn đáng cân nhắc cho {purpose} của bạn!"

    return highlight

# Hàm tìm kiếm sản phẩm theo mã sản phẩm, chương trình khuyến mãi
def search_by_code(code):
    matched_products = df[df['Mã sản phẩm'].str.lower() == code]
    if not matched_products.empty:
        row = matched_products.iloc[0]
        response = (
            f"**{row['Tên sản phẩm']} - {row['Mã sản phẩm']}**\n"
            f"- Giá: {row['Giá (VND)']:,} VND\n"
        )
        return response
    return None

# Hàm trích xuất giá từ câu hỏi
def extract_price(input_text):
    numbers = re.findall(r'\d+', input_text)
    if not numbers:
        return None

    price = 0
    for i, num in enumerate(numbers):
        num = int(num)
        segment = input_text.lower()
        if "triệu" in segment or "tr" in segment:
            price += num * 1000000
        elif "nghìn" in segment or "k" in segment:
            price += num * 1000
        else:
            price += num
    return price

# Hàm xử lý tin nhắn tự nhiên
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_input = update.message.text.lower().strip()
    normalized_input = normalize_string(user_input)
    tz = pytz.timezone('Asia/Ho_Chi_Minh')
    user_input = update.message.text if update.message and update.message.text else ""
    print(f"Received user input: {user_input}")  # Log để debug

    # Lưu tin nhắn người dùng
    await log_user_message(update, context)

    # Từ khóa để nhận diện câu hỏi liên quan đến ngữ cảnh
    context_keywords = ["thì sao", "còn", "thế", "vậy"]

    # Kiểm tra câu hỏi về chương trình khuyến mãi
    if "khuyến mãi" in user_input or "khuyến mại" in user_input or "ưu đãi" in user_input:
        # Kiểm tra nếu hỏi khuyến mãi cho một mã sản phẩm cụ thể
        for code in df['Mã sản phẩm'].str.lower():
            if code in user_input:
                result = search_promotions(product_code=code)
                await send_and_log(update, context, result)
                # Lưu ngữ cảnh
                context.user_data['last_action'] = 'search_promotions'
                context.user_data['last_product_code'] = code
                return
            # Kiểm tra nếu hỏi khuyến mãi cho một danh mục
            for normalized_category, original_category in category_mapping.items():
                if normalized_category in normalized_input:
                    result = search_promotions(category=original_category)
                    await send_and_log(update, context, result)
                    # Lưu ngữ cảnh
                    context.user_data['last_action'] = 'search_promotions'
                    context.user_data['last_category'] = original_category
                    return

            # Kiểm tra nếu hỏi khuyến mãi theo mục đích
            purposes = ["hội nghị", "sự kiện", "quảng cáo", "phòng họp", "livestream", "trung tâm điều hành",
                        "đơn hàng lớn"]
            for purpose in purposes:
                if purpose in user_input:
                    result = search_promotions(purpose=purpose)
                    await send_and_log(update, context, result)
                    # Lưu ngữ cảnh
                    context.user_data['last_action'] = 'search_promotions'
                    context.user_data['last_purpose'] = purpose
                    return
            # Nếu không chỉ định danh mục, sản phẩm, hoặc mục đích, trả về tất cả khuyến mãi
            result = search_promotions()
            await send_and_log(update, context, result)
            # Lưu ngữ cảnh
            context.user_data['last_action'] = 'search_promotions'
            return
        # Kiểm tra nếu câu hỏi có liên quan đến ngữ cảnh trước đó
        if any(keyword in user_input for keyword in context_keywords):
            if 'last_action' in context.user_data and context.user_data['last_action'] == 'search_promotions':
                if 'last_category' in context.user_data:
                    # Hỏi tiếp về danh mục khác
                    for normalized_category, original_category in category_mapping.items():
                        if normalized_category in normalized_input:
                            result = search_promotions(category=original_category)
                            await send_and_log(update, context, result)
                            context.user_data['last_action'] = 'search_promotions'
                            context.user_data['last_category'] = original_category
                            return
                    # Hỏi về các khuyến mãi khác
                    if "khác" in user_input or "nữa" in user_input:
                        result = search_promotions()
                        await send_and_log(update, context, result)
                        context.user_data['last_action'] = 'search_promotions'
                        return
                elif 'last_purpose' in context.user_data:
                    # Hỏi tiếp về mục đích khác
                    purposes = ["hội nghị", "sự kiện", "quảng cáo", "phòng họp", "livestream", "trung tâm điều hành",
                                "đơn hàng lớn"]
                    for purpose in purposes:
                        if purpose in user_input:
                            result = search_promotions(purpose=purpose)
                            await send_and_log(update, context, result)
                            context.user_data['last_action'] = 'search_promotions'
                            context.user_data['last_purpose'] = purpose
                            return
                    # Hỏi về các khuyến mãi khác
                    if "khác" in user_input or "nữa" in user_input:
                        result = search_promotions()
                        await send_and_log(update, context, result)
                        context.user_data['last_action'] = 'search_promotions'
                        return
    # Kiểm tra nếu người dùng chỉ hỏi danh mục (VD: "proav", "videowall")
    for normalized_category, original_category in category_mapping.items():
        if normalized_input == normalized_category:
            result = get_product_codes_by_category(original_category)
            if result:
                await send_and_log(update, context, result)
                # Lưu ngữ cảnh: người dùng hỏi về danh sách sản phẩm của danh mục
                context.user_data['last_action'] = 'list_products'
                context.user_data['last_category'] = original_category
                return

    # Kiểm tra nếu người dùng hỏi về danh sách mã sản phẩm theo danh mục
    for normalized_category, original_category in category_mapping.items():
        if normalized_category in normalized_input and (
                "các loại" in user_input or "loại" in user_input or "cung cấp" in user_input or "mã sản phẩm" in user_input):
            result = get_product_codes_by_category(original_category)
            if result:
                await send_and_log(update, context, result)
                # Lưu ngữ cảnh
                context.user_data['last_action'] = 'list_products'
                context.user_data['last_category'] = original_category
                return

        # Kiểm tra nếu câu hỏi có liên quan đến ngữ cảnh trước đó
        if any(keyword in user_input for keyword in context_keywords):
            # Kiểm tra nếu có ngữ cảnh trước đó
            if 'last_action' in context.user_data and 'last_category' in context.user_data:
                last_action = context.user_data['last_action']
                last_category = context.user_data['last_category']

                # Xử lý câu hỏi "còn loại nào khác không?"
                if "loại nào khác" in user_input or "còn loại nào" in user_input:
                    categories = df['Tên sản phẩm'].unique().tolist()
                    other_categories = [cat for cat in categories if cat != last_category]
                    if other_categories:
                        response = "Ngoài " + last_category + ", KLC còn cung cấp các loại sản phẩm khác:\n" + "\n".join(
                            [f"- {cat}" for cat in other_categories])
                        await send_and_log(update, context, response)
                        context.user_data['last_action'] = 'list_categories'
                        return
                    else:
                        await send_and_log(update, context,
                                           f"Hiện tại chỉ có {last_category} thôi. Bạn có thể xem thêm ở /products nhé!")
                        return

                # Tìm danh mục mới trong câu hỏi hiện tại
                for normalized_category, original_category in category_mapping.items():
                    if normalized_category in normalized_input:
                        # Nếu câu hỏi trước là hỏi về danh sách sản phẩm
                        if last_action == 'list_products':
                            result = get_product_codes_by_category(original_category)
                            if result:
                                await send_and_log(update, context, result)
                                # Cập nhật ngữ cảnh
                                context.user_data['last_action'] = 'list_products'
                                context.user_data['last_category'] = original_category
                                return
                        # Nếu câu hỏi trước là hỏi về giá, chỉ áp dụng nếu câu hỏi hiện tại có từ khóa giá
                        elif last_action == 'search_by_price' and (
                                "giá" in user_input or "dưới" in user_input or "trên" in user_input):
                            price_limit = context.user_data.get('last_price_limit', 0)
                            below = context.user_data.get('last_price_below', True)
                            result = search_by_price(original_category, price_limit, below=below)
                            await send_and_log(update, context, result)
                            # Cập nhật ngữ cảnh
                            context.user_data['last_action'] = 'search_by_price'
                            context.user_data['last_category'] = original_category
                            context.user_data['last_price_limit'] = price_limit
                            context.user_data['last_price_below'] = below
                            return
                        break
            else:
                # Nếu không tìm thấy danh mục rõ ràng, hỏi lại để xác nhận
                for normalized_category, original_category in category_mapping.items():
                    if normalized_category in normalized_input:
                        await send_and_log(update, context,
                                           f"Ý bạn là loại sản phẩm của {original_category} đúng không? "
                                           "Nếu đúng, hãy trả lời 'đúng', nếu không, hãy hỏi lại."
                                           )
                        # Lưu danh mục tiềm năng để xử lý tiếp
                        context.user_data['pending_category'] = original_category
                        context.user_data['pending_action'] = 'list_products'
                        return
                    elif last_action == 'search_by_price' and (
                            "giá" in user_input or "dưới" in user_input or "trên" in user_input):
                        price_limit = context.user_data.get('last_price_limit', 0)
                        below = context.user_data.get('last_price_below', True)
                        await send_and_log(update, context, -
                        f"Ý bạn là {original_category} nào có giá {'dưới' if below else 'trên'} {price_limit:,} VND đúng không? "
                        "Nếu đúng, hãy trả lời 'đúng', nếu không, hãy hỏi lại."
                                           )
                        # Lưu danh mục tiềm năng để xử lý tiếp
                        context.user_data['pending_category'] = original_category
                        context.user_data['pending_action'] = 'list_products'
                        context.user_data['pending_price_limit'] = price_limit
                        context.user_data['pending_price_below'] = below
                        return

        # Kiểm tra nếu người dùng trả lời xác nhận (VD: "đúng")
        if 'pending_action' in context.user_data and user_input in ["đúng", "dung", "ok"]:
            pending_category = context.user_data.get('pending_category')
            pending_action = context.user_data.get('pending_action')
            if pending_action == 'list_products':
                result = get_product_codes_by_category(pending_category)
                if result:
                    await send_and_log(update, context, result)
                    # Cập nhật ngữ cảnh
                    context.user_data['last_action'] = 'list_products'
                    context.user_data['last_category'] = pending_category
            elif pending_action == 'search_by_price':
                price_limit = context.user_data.get('pending_price_limit', 0)
                below = context.user_data.get('pending_price_below', True)
                result = search_by_price(pending_category, price_limit, below=below)
                await send_and_log(update, context, result)

                # Cập nhật ngữ cảnh
                context.user_data['last_action'] = 'search_by_price'
                context.user_data['last_category'] = pending_category
                context.user_data['last_price_limit'] = price_limit
                context.user_data['last_price_below'] = below

            # Xóa trạng thái chờ
            context.user_data.pop('pending_category', None)
            context.user_data.pop('pending_action', None)
            context.user_data.pop('pending_price_limit', None)
            context.user_data.pop('pending_price_below', None)
            return

    # Kiểm tra nếu người dùng hỏi về các loại sản phẩm công ty cung cấp
    if "công ty" in user_input and ("cung cấp" in user_input or "loại" in user_input or "sản phẩm" in user_input):
        categories = df['Tên sản phẩm'].unique().tolist()
        response = "Hiện tại, KLC đang cung cấp các loại sản phẩm:\n" + "\n".join([f"- {cat}" for cat in categories])
        await send_and_log(update, context, response)
        # Lưu ngữ cảnh: người dùng hỏi về danh sách danh mục
        context.user_data['last_action'] = 'list_categories'
        return
    # Kiểm tra nếu người dùng yêu cầu gửi email (bằng từ khóa hoặc /sendemail)
    email_keywords = ["gửi email", "cho tôi email", "gửi thông tin qua email", "gửi mail", "email cho tôi"]
    if user_input.startswith('/sendemail') or any(keyword in user_input for keyword in email_keywords):
        context.user_data['email_request'] = True
        await send_and_log(update, context,
                           "Bạn muốn gửi thông tin gì qua email?\n1. Thông tin sản phẩm\n2. Chương trình khuyến mãi\n3. Lịch sử trò chuyện\n4. Tất cả\nVui lòng chọn số từ 1 đến 4.")
        return
    # Kiểm tra nếu người dùng hỏi về mục đích (VD: "Video Wall nào phù hợp cho hội nghị?")
    purposes = ["hội nghị", "sự kiện", "quảng cáo", "phòng họp", "livestream", "trung tâm điều hành"]

    # Tìm loại sản phẩm (category) từ user_input trước
    matched_category = None
    normalized_input = normalize_string(user_input)

    # Duyệt qua category_mapping để tìm loại sản phẩm phù hợp nhất
    for normalized_category, original_category in category_mapping.items():
        if normalized_category in normalized_input:
            matched_category = (normalized_category, original_category)
            break

    # Nếu tìm thấy loại sản phẩm, tiếp tục kiểm tra mục đích
    if matched_category:
        normalized_category, original_category = matched_category

        # Kiểm tra mục đích (purpose) trong user_input
        for purpose in purposes:
            if purpose in user_input and "phù hợp" in user_input:
                result = search_by_purpose(original_category, purpose)
                await send_and_log(update, context, result)
                # Lưu ngữ cảnh
                context.user_data['last_action'] = 'search_by_purpose'
                context.user_data['last_category'] = original_category
                return

    # Kiểm tra nếu người dùng hỏi về giá (VD: "Video Wall nào có giá dưới 10 triệu?")
    for normalized_category, original_category in category_mapping.items():
        # Kiểm tra câu hỏi về giá (dưới)
        if normalized_category in normalized_input and (
                "giá dưới" in user_input or "dưới" in user_input):
            price_limit = extract_price(user_input)
            if price_limit is not None:
                result = search_by_price(original_category, price_limit, below=True)
                if result:
                    await send_and_log(update, context, result)
                else:
                    await send_and_log(update, context,
                                       f"Xin lỗi, hiện tại không có {original_category} nào dưới {price_limit:,} VND.")
                # Lưu ngữ cảnh đầy đủ
                context.user_data['last_action'] = 'search_by_price'
                context.user_data['last_category'] = original_category
                context.user_data['last_price_limit'] = price_limit
                context.user_data['last_price_below'] = True
                return

        # Kiểm tra câu hỏi về giá (trên)
        if normalized_category in normalized_input and (
                "giá trên" in user_input or "trên" in user_input):
            price_limit = extract_price(user_input)
            if price_limit is not None:
                result = search_by_price(original_category, price_limit, below=False)
                if result:
                    await send_and_log(update, context, result)
                else:
                    await send_and_log(update, context,
                                       f"Xin lỗi, hiện tại không có {original_category} nào trên {price_limit:,} VND.")
                # Lưu ngữ cảnh đầy đủ
                context.user_data['last_action'] = 'search_by_price'
                context.user_data['last_category'] = original_category
                context.user_data['last_price_limit'] = price_limit
                context.user_data['last_price_below'] = False
                return

        # Xử lý câu hỏi chỉ đề cập danh mục (dựa trên ngữ cảnh giá trước đó)
        if normalized_category in normalized_input and (
                "giá" in user_input or "dưới" in user_input or "trên" in user_input):
            last_action = context.user_data.get('last_action')
            if last_action == 'search_by_price':
                price_limit = context.user_data.get('last_price_limit', 0)
                below = context.user_data.get('last_price_below', True)
                result = search_by_price(original_category, price_limit, below=below)
                if result:
                    await send_and_log(update, context, result)
                else:
                    await send_and_log(update, context,
                                       f"Xin lỗi, hiện tại không có {original_category} nào {'dưới' if below else 'trên'} {price_limit:,} VND."
                                       )
                # Cập nhật ngữ cảnh
                context.user_data['last_action'] = 'search_by_price'
                context.user_data['last_category'] = original_category
                context.user_data['last_price_limit'] = price_limit
                context.user_data['last_price_below'] = below
                return

    # Kiểm tra nếu người dùng hỏi về giá của một sản phẩm cụ thể
    if "giá" in user_input or "bao nhiêu" in user_input:
        for code in df['Mã sản phẩm'].str.lower():
            if code in user_input:
                result = search_by_code(code)
                if result:
                    await send_and_log(update, context, result)
                    # Lưu ngữ cảnh
                    context.user_data['last_action'] = 'search_by_code'
                    return
        await send_and_log(update, context,
                           "Vui lòng cung cấp mã sản phẩm để tôi kiểm tra giá, ví dụ: 'Giá của AVE-HU50D là bao nhiêu?'")
        return

    # Kiểm tra nếu người dùng hỏi về yếu tố trong mô tả
    description_keywords = [
        "4k", "hdmi", "displayport", "led", "16 cổng", "hội nghị", "sự kiện",
        "trung tâm điều hành", "quảng cáo", "phòng họp", "livestream", "chống chói"
    ]

    # Tìm danh mục trước khi tìm từ khóa trong mô tả
    matched_category = None
    for normalized_category, original_category in category_mapping.items():
        if normalized_category in normalized_input:
            matched_category = (normalized_category, original_category)
            break

    for keyword in description_keywords:
        if keyword in user_input.lower():
            # Nếu có danh mục, lọc theo danh mục trước
            if matched_category:
                normalized_category, original_category = matched_category
                matched_products = df[
                    (df['Tên sản phẩm'] == original_category) &
                    (df['Mô tả'].str.lower().str.contains(keyword, na=False))
                    ]
            else:
                # Nếu không có danh mục, tìm trong toàn bộ dữ liệu
                matched_products = df[df['Mô tả'].str.lower().str.contains(keyword, na=False)]

            if not matched_products.empty:
                response = f"Thiết bị hỗ trợ {keyword} phù hợp cho bạn đây!\n\n"

                # Danh sách để lưu trữ thông tin sản phẩm
                product_highlights = []

                for _, row in matched_products.iterrows():
                    product_code = row['Mã sản phẩm']
                    description = row['Mô tả']

                    # Tóm tắt và diễn đạt lại mô tả
                    highlight = generate_highlight(description, keyword, product_code, style="practical_friendly")
                    product_highlights.append((product_code, highlight))

                # Thêm thông tin sản phẩm vào response
                for product_code, highlight in product_highlights:
                    response += (
                        f"**{product_code}**  \n"
                        f"- **Điểm nổi bật**: {highlight}  \n\n"
                    )

                # Thêm phần gợi ý với phong cách thực tế, thân thiện
                if len(product_highlights) > 1:
                    response += "**Gợi ý chọn thiết bị**:  \n"
                    if keyword.lower() == "4k":
                        for product_code, highlight in product_highlights:
                            desc_lower = highlight.lower()
                            if "trung tâm điều hành" in desc_lower or "phòng họp lớn" in desc_lower:
                                response += f"- Dùng cho công việc chuyên nghiệp (trung tâm điều hành, phòng họp): {product_code} là lựa chọn tốt nhờ khả năng quản lý nhiều máy chủ.  \n"
                            if "livestream" in desc_lower or "chuyển đổi nhanh" in desc_lower:
                                response += f"- Cần livestream hoặc chuyển đổi nhanh: {product_code} sẽ đáp ứng tốt với tính năng phím nóng và hỗ trợ 4K mượt mà.  \n"
                            if "quảng cáo" in desc_lower or "sự kiện lớn" in desc_lower:
                                response += f"- Dùng cho trình chiếu (quảng cáo, sự kiện): {product_code} là lựa chọn lý tưởng với hình ảnh 4K sắc nét.  \n"
                    else:
                        response += (
                            f"- Dùng cho {keyword} lớn, cần thiết bị mạnh: Chọn sản phẩm có khả năng kết nối nhiều thiết bị và hiệu suất cao.  \n"
                            f"- Dùng cho {keyword} vừa và nhỏ, cần linh hoạt: Chọn sản phẩm dễ sử dụng và đa năng.  \n"
                        )
                    response += f"- **Lưu ý**: Hãy xem bạn cần {keyword} cho mục đích gì để chọn thiết bị phù hợp nhé!  "
                else:
                    response += f"**Lưu ý**: Hãy xem bạn cần {keyword} cho mục đích gì để tận dụng thiết bị này tốt nhất nhé!  "

                await send_and_log(update, context, response)
                # Lưu ngữ cảnh
                context.user_data['last_action'] = 'search_by_description'
                return
            # Nếu không tìm thấy gì phù hợp
            categories = df['Tên sản phẩm'].unique().tolist()
            response = (
                    "Mình chưa tìm thấy thiết bị phù hợp với yêu cầu của bạn. "
                    "Bạn có thể thử hỏi về các danh mục sau:\n" + "\n".join([f"- {cat}" for cat in categories]) +
                    "\nHoặc xem gợi ý ở /help nhé!"
            )

    # Kiểm tra trạng thái gửi email
    if 'email_step' in context.user_data:
        step = context.user_data['email_step']
        if step == 'ask_email':
            if user_input.lower() == 'hủy':
                context.user_data.pop('email_step', None)
                context.user_data.pop('email_choice', None)
                context.user_data.pop('email_address', None)
                await send_and_log(update, context, "Đã hủy gửi email. Bạn cần hỗ trợ gì tiếp theo?")
                return
            if re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', user_input):
                context.user_data['email_address'] = user_input
                context.user_data['email_step'] = 'confirm_content'
                choice = context.user_data.get('email_choice', '1')
                content_desc = {
                    '1': 'thông tin sản phẩm',
                    '2': 'chương trình khuyến mãi',
                    '3': 'lịch sử trò chuyện',
                    '4': 'tất cả các mục trên'
                }
                await send_and_log(update, context,
                                   f"Tôi sẽ gửi {content_desc[choice]} đến {user_input}. Trả lời 'gửi' để xác nhận hoặc 'hủy' để thoát.")
            else:
                await send_and_log(update, context,
                                   "Địa chỉ email không hợp lệ. Vui lòng nhập lại (VD: example@domain.com) hoặc 'hủy' để thoát.")
            return
        elif step == 'confirm_content':
            if user_input.lower() == 'gửi':
                email_address = context.user_data.get('email_address')
                choice = context.user_data.get('email_choice', '1')

                # Tạo nội dung email
                products_html = df[['Mã sản phẩm', 'Tên sản phẩm', 'Giá (VND)']].to_html(index=False,
                                                                                         classes='dataframe')
                promotions_html = df_promo[
                    ['Mã khuyến mãi', 'Loại khuyến mãi', 'Giá trị', 'Áp dụng cho', 'Ngày bắt đầu',
                     'Ngày kết thúc']].to_html(index=False, classes='dataframe')

                today = datetime.now(tz).date()
                convo_html = "<h2>Lịch sử trò chuyện hôm nay</h2><ul>"
                print(f"Conversation history: {context.user_data.get('conversation', [])}")  # Log để debug
                for entry in context.user_data.get('conversation', []):
                    if len(entry) == 3:
                        sender, message, timestamp = entry
                        if timestamp.date() == today:
                            name = "Người dùng" if sender == 'user' else "Bot"
                            time_str = timestamp.strftime('%H:%M')
                            convo_html += f"<li><strong>{name} ({time_str}):</strong> {message}</li>"
                convo_html += "</ul>"
                if convo_html == "<h2>Lịch sử trò chuyện hôm nay</h2><ul></ul>":
                    convo_html = "<h2>Lịch sử trò chuyện hôm nay</h2><p>Không có tin nhắn nào trong hôm nay.</p>"

                html_body = f"""
                    <html>
                    <head><meta charset="utf-8"></head>
                    <body>
                        <h1>Thông tin từ KLC</h1>
                        {'<h2>Thông tin sản phẩm</h2>' + products_html if choice in ['1', '4'] else ''}
                        {'<h2>Chương trình khuyến mãi</h2>' + promotions_html if choice in ['2', '4'] else ''}
                        {convo_html if choice in ['3', '4'] else ''}
                        <hr>
                        <p>Cảm ơn bạn đã sử dụng dịch vụ của KLC.</p>
                    </body>
                    </html>
                    """

                print(f"Sending email to {email_address} with content: {html_body}")  # Log để debug
                success = await send_email(email_address, "Thông tin từ KLC ChatBot", html_body)
                if success:
                    await send_and_log(update, context, f"Email đã được gửi đến {email_address} thành công!")
                else:
                    await send_and_log(update, context, "Có lỗi xảy ra khi gửi email. Vui lòng thử lại sau.")
                # Xóa trạng thái email
                context.user_data.pop('email_step', None)
                context.user_data.pop('email_choice', None)
                context.user_data.pop('email_address', None)
            else:
                await send_and_log(update, context, "Bạn đã hủy gửi email.")
                context.user_data.pop('email_step', None)
                context.user_data.pop('email_choice', None)
                context.user_data.pop('email_address', None)
            return

    # Xử lý yêu cầu gửi email
    if user_input in ['1', '2', '3', '4']:
        context.user_data['email_request'] = True
        context.user_data['email_choice'] = user_input
        await send_and_log(update, context,
                           "Vui lòng cung cấp địa chỉ email của bạn (VD: example@domain.com). Hoặc trả lời 'hủy' để thoát.")
        context.user_data['email_step'] = 'ask_email'
        return
    elif context.user_data.get('email_request', False):
        await send_and_log(update, context, "Lựa chọn không hợp lệ. Vui lòng chọn số từ 1 đến 4.")
        context.user_data['email_request'] = False
        return

    # Xử lý các tin nhắn khác (nếu có)
    await send_and_log(update, context, "Vui lòng chọn một tùy chọn (1-4) để nhận thông tin qua email.")


# Hàm trích xuất sản phẩm/danh mục hoặc mã khuyến mãi từ câu hỏi
def extract_product_or_promo(user_input, df_products=None, df_promo=None):
    products = df_promo['Áp dụng cho'].str.lower().tolist()
    promos = df_promo['Mã khuyến mãi'].str.lower().tolist()

    if df_products is not None:
        products += df_products['Tên'].str.lower().tolist()

    for item in products:
        if item.lower() in user_input.lower():
            return item, "product"
    for promo in promos:
        if promo.lower() in user_input.lower():
            return promo, "promo"
    return None, None

# Hàm lấy điều kiện áp dụng theo sản phẩm/danh mục
def get_promo_conditions(product_or_category):
    valid_promos = get_valid_promotions()
    # Tìm kiếm sản phẩm/danh mục chính xác hơn
    promo = valid_promos[valid_promos['Áp dụng cho'].str.lower().str.contains(product_or_category.lower(), na=False)]
    if not promo.empty:
        response = f"Ưu đãi và điều kiện áp dụng cho {product_or_category}:\n"
        for _, p in promo.iterrows():
            condition = p['Điều kiện áp dụng'] if p['Điều kiện áp dụng'] != "<unset>" else "Không có điều kiện cụ thể."
            response += f"- {p['Loại khuyến mãi']} ({p['Giá trị']}): {condition}\n"
        return response
    return f"Không tìm thấy ưu đãi nào cho {product_or_category} vào thời điểm hiện tại."


# Hàm lấy điều kiện áp dụng theo mã khuyến mãi
def get_promo_conditions_by_code(promo_code):
    valid_promos = get_valid_promotions()
    promo = valid_promos[valid_promos['Mã khuyến mãi'].str.lower() == promo_code.lower()]
    if not promo.empty:
        promo = promo.iloc[0]
        condition = promo['Điều kiện áp dụng'] if promo[
                                                      'Điều kiện áp dụng'] != "<unset>" else "Không có điều kiện cụ thể."
        response = f"Ưu đãi và điều kiện áp dụng cho mã {promo['Mã khuyến mãi']}:\n"
        response += f"- {promo['Loại khuyến mãi']} ({promo['Giá trị']}) cho {promo['Áp dụng cho']}: {condition}"
        return response
    return f"Không tìm thấy chương trình khuyến mãi {promo_code}."

# Hàm check_promo_eligibility cải tiến
def check_promo_eligibility(product_name, quantity):
    valid_promos = get_valid_promotions()
    promo = valid_promos[valid_promos['Áp dụng cho'].str.lower().str.contains(product_name.lower(), na=False)]

    if promo.empty:
        return f"Sản phẩm {product_name} hiện không có ưu đãi."

    response = f"Kết quả kiểm tra mua {quantity} sản phẩm {product_name}:\n"
    for _, p in promo.iterrows():
        condition = p['Điều kiện áp dụng'] if p['Điều kiện áp dụng'] != "<unset>" else ""
        if "Mua từ" in condition:
            try:
                required_quantity = int(condition.split("Mua từ")[1].split("sản phẩm")[0].strip())
                if quantity >= required_quantity:
                    response += f"- Đủ điều kiện nhận: {p['Loại khuyến mãi']} ({p['Giá trị']}).\n"
                else:
                    response += f"- Cần mua ít nhất {required_quantity} sản phẩm để nhận: {p['Loại khuyến mãi']} ({p['Giá trị']}). Bạn chỉ mua {quantity} sản phẩm.\n"
            except:
                response += f"- {p['Loại khuyến mãi']} ({p['Giá trị']}): {condition}\n"
        else:
            response += f"- Nhận được: {p['Loại khuyến mãi']} ({p['Giá trị']}). Điều kiện: {condition or 'Không có điều kiện cụ thể.'}\n"

    response += "\nBạn có muốn biết thêm về ưu đãi cho sản phẩm khác không?"
    return response

def process_user_input(user_input, conversation_context=None):
    promo_keywords = ["khuyến mãi", "ưu đãi", "giảm giá", "promo", "promotion"]
    condition_keywords = ["điều kiện", "cần làm gì", "yêu cầu", "để được", "condition", "requirement", "áp dụng",
                          "thế nào", "làm sao"]

    # Khởi tạo ngữ cảnh nếu chưa có
    if conversation_context is None:
        conversation_context = {"last_product": None, "last_response_type": None}

    # Hỏi chung về khuyến mãi
    if any(keyword in user_input.lower() for keyword in
           promo_keywords) and "gì" in user_input.lower() and "mua" not in user_input.lower():
        valid_promos = get_valid_promotions()
        if not valid_promos.empty:
            response = "Hiện tại có các ưu đãi sau:\n"
            for _, p in valid_promos.iterrows():
                response += f"- {p['Áp dụng cho']}: {p['Loại khuyến mãi']} ({p['Giá trị']}), đến {p['Ngày kết thúc']}.\n"
            response += "\nHỏi cụ thể để biết thêm chi tiết nhé! Ví dụ: 'Ưu đãi cho Video Wall'."
            conversation_context["last_response_type"] = "list"
            return response, conversation_context
        return "Hiện không có ưu đãi nào.", conversation_context

    # Trích xuất sản phẩm hoặc mã khuyến mãi
    item, item_type = extract_product_or_promo(user_input, df_promo=df_promo)

    # Cập nhật ngữ cảnh
    if item and item_type == "product":
        conversation_context["last_product"] = item

    # Xử lý câu hỏi về số lượng
    if "mua" in user_input.lower():
        try:
            quantity_match = re.search(r'mua\s+(\d+)\s*(?:sản phẩm)?', user_input.lower())
            if quantity_match:
                quantity = int(quantity_match.group(1))
                if item and item_type == "product":
                    response = check_promo_eligibility(item, quantity)
                    conversation_context["last_response_type"] = "quantity"
                    return response, conversation_context
                return "Vui lòng chỉ rõ sản phẩm bạn muốn mua, ví dụ: 'Mua 2 Video Wall có khuyến mãi gì?'", conversation_context
        except:
            pass

    # Xử lý câu hỏi về điều kiện hoặc ưu đãi cụ thể
    if item:
        if item_type == "product":
            # Nếu hỏi về điều kiện
            if any(keyword in user_input.lower() for keyword in condition_keywords):
                response = get_promo_conditions(item)
                response += "\nBạn có muốn kiểm tra số lượng mua để nhận ưu đãi không? Ví dụ: 'Mua 2 Video Wall có khuyến mãi gì?'"
                conversation_context["last_response_type"] = "condition"
                return response, conversation_context
            # Nếu hỏi chung về ưu đãi (ví dụ: "Ưu đãi cho video wall")
            # Kiểm tra ngữ cảnh để tránh lặp lại
            if conversation_context["last_product"] == item and conversation_context["last_response_type"] == "list":
                response = f"Bạn vừa hỏi về ưu đãi cho {item}. Đây là thông tin chi tiết:\n"
            else:
                response = f"Ưu đãi cho {item}:\n"
            promo = get_valid_promotions()[
                get_valid_promotions()['Áp dụng cho'].str.lower().str.contains(item.lower(), na=False)]
            for _, p in promo.iterrows():
                condition = p['Điều kiện áp dụng'] if p[
                                                          'Điều kiện áp dụng'] != "<unset>" else "Không có điều kiện cụ thể."
                response += f"- {p['Loại khuyến mãi']} ({p['Giá trị']}): {condition}\n"
            response += "\nHỏi thêm về số lượng nếu bạn muốn nhé! Ví dụ: 'Mua 2 Video Wall có khuyến mãi gì?'"
            conversation_context["last_response_type"] = "specific"
            return response, conversation_context
        elif item_type == "promo":
            response = get_promo_conditions_by_code(item)
            conversation_context["last_response_type"] = "promo"
            return response, conversation_context

    # Nếu không tìm thấy sản phẩm/mã, gợi ý các danh mục
    if any(keyword in user_input.lower() for keyword in promo_keywords):
        valid_promos = get_valid_promotions()
        categories = valid_promos['Áp dụng cho'].unique().tolist()
        response = "Không tìm thấy sản phẩm hoặc mã khuyến mãi phù hợp. Bạn có thể thử hỏi về:\n"
        for cat in categories:
            response += f"- {cat}\n"
        response += "Hoặc xem gợi ý ở /help nhé!"
        conversation_context["last_response_type"] = "not_found"
        return response, conversation_context

    return "Tôi chưa hiểu câu hỏi của bạn. Bạn có thể hỏi lại không?", conversation_context

def get_valid_promotions():
    current_date = datetime.now().date()  # Chỉ lấy ngày, bỏ thời gian
    print(f"Ngày hiện tại (chỉ ngày): {current_date}")

    # Chuyển đổi cột Ngày bắt đầu và Ngày kết thúc thành chỉ ngày
    valid_promos = df_promo[
        (df_promo['Ngày bắt đầu'].notna()) &
        (df_promo['Ngày kết thúc'].notna()) &
        (df_promo['Ngày bắt đầu'].dt.date <= current_date) &
        (df_promo['Ngày kết thúc'].dt.date >= current_date)
        ]
    print(f"Các khuyến mãi hợp lệ:\n{valid_promos}")
    return valid_promos

def search_promotions(category=None, product_code=None, purpose=None):
    valid_promos = get_valid_promotions()
    if valid_promos.empty:
        return "Hiện tại không có chương trình khuyến mãi nào đang diễn ra."

    response = ""
    matched_promos = pd.DataFrame()

    if product_code:
        # Tìm khuyến mãi theo mã sản phẩm
        matched_promos = valid_promos[valid_promos['Áp dụng cho'].str.lower() == product_code.lower()]
        if matched_promos.empty:
            # Kiểm tra nếu mã sản phẩm thuộc danh mục có khuyến mãi
            product = df[df['Mã sản phẩm'].str.lower() == product_code.lower()]
            if not product.empty:
                category = product.iloc[0]['Tên sản phẩm']
                matched_promos = valid_promos[valid_promos['Áp dụng cho'] == category]
                if matched_promos.empty:
                    return f"Hiện tại không có chương trình khuyến mãi cho {product_code}."

    elif category:
        # Tìm khuyến mãi theo danh mục
        normalized_category = normalize_string(category)
        original_category = category_mapping.get(normalized_category)
        if not original_category:
            return f"Không tìm thấy danh mục {category}."
        matched_promos = valid_promos[valid_promos['Áp dụng cho'] == original_category]

    elif purpose:
        # Tìm khuyến mãi theo mục đích sử dụng
        matched_promos = valid_promos[
            valid_promos['Điều kiện áp dụng'].str.lower().str.contains(purpose.lower(), na=False)]

    else:
        # Trả về tất cả khuyến mãi
        matched_promos = valid_promos

    if not matched_promos.empty:
        response = "Chương trình khuyến mãi hiện tại:\n\n"
        for _, promo in matched_promos.iterrows():
            response += (
                f"- **{promo['Mã khuyến mãi']}**: {promo['Loại khuyến mãi']} ({promo['Giá trị']})\n"
                f"  + Áp dụng cho: {promo['Áp dụng cho']}\n"
                f"  + Hiệu lực: {promo['Ngày bắt đầu'].strftime('%d/%m/%Y')} - {promo['Ngày kết thúc'].strftime('%d/%m/%Y')}\n"
            )
            if pd.notna(promo['Điều kiện áp dụng']) and promo['Điều kiện áp dụng'].strip():
                response += f"  + Điều kiện: {promo['Điều kiện áp dụng']}\n"
            response += "\n"
    else:
        response = "Hiện tại không có chương trình khuyến mãi phù hợp với yêu cầu của bạn.\n"
        # Gợi ý các khuyến mãi khác
        if not valid_promos.empty:
            response += "Bạn có thể quan tâm đến các chương trình khác:\n"
            for _, promo in valid_promos.head(2).iterrows():  # Giới hạn gợi ý 2 chương trình
                response += f"- {promo['Loại khuyến mãi']} ({promo['Giá trị']}) cho {promo['Áp dụng cho']}, đến {promo['Ngày kết thúc'].strftime('%d/%m/%Y')}.\n"

    return response


# Hàm main để khởi động bot
def main():
    application = Application.builder().token(TOKEN).build()

    # Đăng ký các lệnh
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("products", products))
    application.add_handler(CommandHandler("sendemail", sendemail))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Chạy bot
    print("Bot đang chạy...")
    application.run_polling()


if __name__ == '__main__':
    main()

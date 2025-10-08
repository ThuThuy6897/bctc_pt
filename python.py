import streamlit as st
import pandas as pd
from google import genai
from google.genai.errors import APIError
from google.genai import types # Thêm import types cho việc quản lý lịch sử

# --- Cấu hình Trang Streamlit ---
st.set_page_config(
    page_title="App Phân Tích Báo Cáo Tài Chính",
    layout="wide"
)

st.title("Ứng dụng Phân Tích Báo Cáo Tài Chính 📊")

# --- GLOBAL CONFIG & GEMINI CLIENT SETUP (Phần đã thêm) ---
# Tải API Key và Khởi tạo Client
try:
    # Lấy khóa API từ Streamlit Secrets
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except (KeyError, AttributeError):
    st.error("Lỗi: Không tìm thấy GEMINI_API_KEY trong Streamlit Secrets.")
    st.info("Vui lòng thêm GEMINI_API_KEY vào tệp secrets.toml.")
    st.stop() # Dừng ứng dụng nếu không có key

# Khởi tạo client và mô hình
client = genai.Client(api_key=GEMINI_API_KEY)
MODEL_NAME = 'gemini-2.5-flash'

# Khởi tạo lịch sử trò chuyện (Chat History) trong Streamlit Session State
if "chat" not in st.session_state:
    try:
        # Tạo một phiên trò chuyện mới với mô hình
        st.session_state.chat = client.chats.create(model=MODEL_NAME)
    except Exception as e:
        # Ghi nhận lỗi nhưng không dừng app, để người dùng vẫn dùng được phần Phân Tích
        st.error(f"Lỗi khi khởi tạo mô hình chat: {e}") 

# --- Hàm tính toán chính (Sử dụng Caching để Tối ưu hiệu suất) ---
@st.cache_data
def process_financial_data(df):
    """Thực hiện các phép tính Tăng trưởng và Tỷ trọng."""
    
    # Đảm bảo các giá trị là số để tính toán
    numeric_cols = ['Năm trước', 'Năm sau']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # 1. Tính Tốc độ Tăng trưởng
    # Dùng .replace(0, 1e-9) cho Series Pandas để tránh lỗi chia cho 0
    df['Tốc độ tăng trưởng (%)'] = (
        (df['Năm sau'] - df['Năm trước']) / df['Năm trước'].replace(0, 1e-9)
    ) * 100

    # 2. Tính Tỷ trọng theo Tổng Tài sản
    # Lọc chỉ tiêu "TỔNG CỘNG TÀI SẢN"
    tong_tai_san_row = df[df['Chỉ tiêu'].str.contains('TỔNG CỘNG TÀI SẢN', case=False, na=False)]
    
    if tong_tai_san_row.empty:
        raise ValueError("Không tìm thấy chỉ tiêu 'TỔNG CỘNG TÀI SẢN'.")

    tong_tai_san_N_1 = tong_tai_san_row['Năm trước'].iloc[0]
    tong_tai_san_N = tong_tai_san_row['Năm sau'].iloc[0]

    # ******************************* PHẦN SỬA LỖI BẮT ĐẦU *******************************
    # Lỗi xảy ra khi dùng .replace() trên giá trị đơn lẻ (numpy.int64).
    # Sử dụng điều kiện ternary để xử lý giá trị 0 thủ công cho mẫu số.
    
    divisor_N_1 = tong_tai_san_N_1 if tong_tai_san_N_1 != 0 else 1e-9
    divisor_N = tong_tai_san_N if tong_tai_san_N != 0 else 1e-9

    # Tính tỷ trọng với mẫu số đã được xử lý
    df['Tỷ trọng Năm trước (%)'] = (df['Năm trước'] / divisor_N_1) * 100
    df['Tỷ trọng Năm sau (%)'] = (df['Năm sau'] / divisor_N) * 100
    # ******************************* PHẦN SỬA LỖI KẾT THÚC *******************************
    
    return df

# --- Hàm gọi API Gemini (Đã điều chỉnh để dùng client toàn cục) ---
def get_ai_analysis(data_for_ai):
    """Gửi dữ liệu phân tích đến Gemini API và nhận nhận xét. Sử dụng client đã khởi tạo."""
    try:
        # Sử dụng client toàn cục đã được khởi tạo ở đầu script
        global client 
        global MODEL_NAME
        
        prompt = f"""
        Bạn là một chuyên gia phân tích tài chính chuyên nghiệp. Dựa trên các chỉ số tài chính sau, hãy đưa ra một nhận xét khách quan, ngắn gọn (khoảng 3-4 đoạn) về tình hình tài chính của doanh nghiệp. Đánh giá tập trung vào tốc độ tăng trưởng, thay đổi cơ cấu tài sản và khả năng thanh toán hiện hành.
        
        Dữ liệu thô và chỉ số:
        {data_for_ai}
        """

        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=prompt
        )
        return response.text

    except APIError as e:
        return f"Lỗi gọi Gemini API: Vui lòng kiểm tra Khóa API hoặc giới hạn sử dụng. Chi tiết lỗi: {e}"
    except Exception as e:
        return f"Đã xảy ra lỗi không xác định: {e}"


# --- Chức năng 1: Tải File (GIỮ NGUYÊN) ---
uploaded_file = st.file_uploader(
    "1. Tải file Excel Báo cáo Tài chính (Chỉ tiêu | Năm trước | Năm sau)",
    type=['xlsx', 'xls']
)

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        # Tiền xử lý: Đảm bảo chỉ có 3 cột quan trọng
        df_raw.columns = ['Chỉ tiêu', 'Năm trước', 'Năm sau']
        
        # Xử lý dữ liệu
        df_processed = process_financial_data(df_raw.copy())

        if df_processed is not None:
            
            # --- Chức năng 2 & 3: Hiển thị Kết quả (GIỮ NGUYÊN) ---
            st.subheader("2. Tốc độ Tăng trưởng & 3. Tỷ trọng Cơ cấu Tài sản")
            st.dataframe(df_processed.style.format({
                'Năm trước': '{:,.0f}',
                'Năm sau': '{:,.0f}',
                'Tốc độ tăng trưởng (%)': '{:.2f}%',
                'Tỷ trọng Năm trước (%)': '{:.2f}%',
                'Tỷ trọng Năm sau (%)': '{:.2f}%'
            }), use_container_width=True)
            
            # --- Chức năng 4: Tính Chỉ số Tài chính (GIỮ NGUYÊN) ---
            st.subheader("4. Các Chỉ số Tài chính Cơ bản")
            
            try:
                # Lọc giá trị cho Chỉ số Thanh toán Hiện hành (Ví dụ)
                
                # Lấy Tài sản ngắn hạn
                tsnh_n = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Năm sau'].iloc[0]
                tsnh_n_1 = df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Năm trước'].iloc[0]

                # Lấy Nợ ngắn hạn (Dùng giá trị giả định hoặc lọc từ file nếu có)
                # **LƯU Ý: Thay thế logic sau nếu bạn có Nợ Ngắn Hạn trong file**
                no_ngan_han_N = df_processed[df_processed['Chỉ tiêu'].str.contains('NỢ NGẮN HẠN', case=False, na=False)]['Năm sau'].iloc[0]  
                no_ngan_han_N_1 = df_processed[df_processed['Chỉ tiêu'].str.contains('NỢ NGẮN HẠN', case=False, na=False)]['Năm trước'].iloc[0]

                # Tính toán
                # Xử lý lỗi chia cho 0
                thanh_toan_hien_hanh_N = tsnh_n / (no_ngan_han_N if no_ngan_han_N != 0 else 1e-9)
                thanh_toan_hien_hanh_N_1 = tsnh_n_1 / (no_ngan_han_N_1 if no_ngan_han_N_1 != 0 else 1e-9)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric(
                        label="Chỉ số Thanh toán Hiện hành (Năm trước)",
                        value=f"{thanh_toan_hien_hanh_N_1:.2f} lần"
                    )
                with col2:
                    st.metric(
                        label="Chỉ số Thanh toán Hiện hành (Năm sau)",
                        value=f"{thanh_toan_hien_hanh_N:.2f} lần",
                        delta=f"{thanh_toan_hien_hanh_N - thanh_toan_hien_hanh_N_1:.2f}"
                    )
                    
            except IndexError:
                st.warning("Thiếu chỉ tiêu 'TÀI SẢN NGẮN HẠN' hoặc 'NỢ NGẮN HẠN' để tính chỉ số.")
                thanh_toan_hien_hanh_N = "N/A" # Dùng để tránh lỗi ở Chức năng 5
                thanh_toan_hien_hanh_N_1 = "N/A"
                
            # --- Chức năng 5: Nhận xét AI (Đã điều chỉnh gọi hàm) ---
            st.subheader("5. Nhận xét Tình hình Tài chính (AI)")
            
            # Chuẩn bị dữ liệu để gửi cho AI
            data_for_ai = pd.DataFrame({
                'Chỉ tiêu': [
                    'Toàn bộ Bảng phân tích (dữ liệu thô)', 
                    'Tăng trưởng Tài sản ngắn hạn (%)', 
                    'Thanh toán hiện hành (N-1)', 
                    'Thanh toán hiện hành (N)'
                ],
                'Giá trị': [
                    df_processed.to_markdown(index=False),
                    f"{df_processed[df_processed['Chỉ tiêu'].str.contains('TÀI SẢN NGẮN HẠN', case=False, na=False)]['Tốc độ tăng trưởng (%)'].iloc[0]:.2f}%" if not isinstance(thanh_toan_hien_hanh_N, str) else "N/A", 
                    f"{thanh_toan_hien_hanh_N_1}", 
                    f"{thanh_toan_hien_hanh_N}"
                ]
            }).to_markdown(index=False) 

            if st.button("Yêu cầu AI Phân tích"):
                
                if GEMINI_API_KEY: # Kiểm tra key đã được load ở trên chưa
                    with st.spinner('Đang gửi dữ liệu và chờ Gemini phân tích...'):
                        # GỌI HÀM ĐÃ ĐƯỢNG ĐIỀU CHỈNH (bỏ tham số api_key)
                        ai_result = get_ai_analysis(data_for_ai)
                        st.markdown("**Kết quả Phân tích từ Gemini AI:**")
                        st.info(ai_result)
                else:
                    st.error("Lỗi: Không tìm thấy Khóa API. Vui lòng cấu hình Khóa 'GEMINI_API_KEY' trong Streamlit Secrets.")

    except ValueError as ve:
        st.error(f"Lỗi cấu trúc dữ liệu: {ve}")
    except Exception as e:
        st.error(f"Có lỗi xảy ra khi đọc hoặc xử lý file: {e}. Vui lòng kiểm tra định dạng file.")

else:
    st.info("Vui lòng tải lên file Excel để bắt đầu phân tích.")

# --- PHẦN ĐÃ THÊM: KHUNG CHAT RIÊNG VỚI GEMINI ---
# Sử dụng st.expander để tạo một khung chat gọn gàng, không ảnh hưởng đến luồng chính
st.markdown("---")
with st.expander("🤖 Chat hỏi đáp về Phân tích Tài chính với Gemini (Hội thoại riêng)", expanded=True):
    
    # Kiểm tra xem chat đã được khởi tạo thành công chưa
    if "chat" in st.session_state:
        
        # Hiển thị lịch sử trò chuyện
        # Sử dụng try-except để phòng trường hợp lỗi đọc parts[0].text
        for message in st.session_state.chat.get_history():
            role = "user" if message.role == "user" else "assistant"
            try:
                content = message.parts[0].text
                with st.chat_message(role):
                    st.markdown(content)
            except Exception:
                # Bỏ qua tin nhắn bị lỗi
                continue

        # Xử lý input mới từ người dùng
        if prompt := st.chat_input("Hỏi về các chỉ số, khái niệm tài chính hoặc dữ liệu đã tải lên..."):
            # 1. Thêm tin nhắn của người dùng vào giao diện
            with st.chat_message("user"):
                st.markdown(prompt)

            # 2. Gửi tin nhắn đến mô hình Gemini và nhận phản hồi
            try:
                # Gửi tin nhắn và nhận phản hồi
                response = st.session_state.chat.send_message(prompt)

                # 3. Thêm phản hồi của Gemini vào giao diện
                with st.chat_message("assistant"):
                    st.markdown(response.text)

            except APIError as e:
                 st.error(f"Lỗi gọi API trong khung chat: {e}. Vui lòng kiểm tra lại GEMINI_API_KEY.")
            except Exception as e:
                st.error(f"Đã xảy ra lỗi không xác định trong khung chat: {e}")
    else:
        st.warning("Không thể khởi tạo khung chat. Vui lòng đảm bảo GEMINI_API_KEY đã được cấu hình chính xác.")


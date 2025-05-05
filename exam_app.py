import streamlit as st
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt
import io
import tempfile
import os

# Set page title and configuration
st.set_page_config(
    page_title="Phần Mềm Tạo Đề Thi",
    page_icon="📝",
    layout="wide"
)

# Title and description
st.title("Phần Mềm Tạo Đề Thi")
st.markdown("""
Ứng dụng này cho phép bạn:
1. Tải lên thư viện câu hỏi Excel của bạn
2. Chọn ngẫu nhiên một số câu hỏi
3. Xuất những câu hỏi đó ra tài liệu MS Word
""")

# Initialize session state variables if they don't exist
if 'question_library' not in st.session_state:
    st.session_state.question_library = None
if 'file_uploaded' not in st.session_state:
    st.session_state.file_uploaded = False
if 'file_valid' not in st.session_state:
    st.session_state.file_valid = False

# File upload section
st.header("Tải Lên Thư Viện Câu Hỏi")
st.markdown("Tải lên một tệp Excel (.xlsx hoặc .xls) chứa thư viện câu hỏi của bạn. Tệp phải có ít nhất 10 cột. Các cột theo thứ tự là: 'Ma cau', 'Cau hoi', 'Tra loi 1', 'Tra loi 2', 'Tra loi 3', 'Tra loi 4', 'Dap an dung', 'Bai', 'Phan', 'Do kho'.")

uploaded_file = st.file_uploader("Chọn tệp Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Attempt to read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Define the column names
        column_names = [
            "Ma cau", "Cau hoi", "Tra loi 1", "Tra loi 2", "Tra loi 3", 
            "Tra loi 4", "Dap an dung", "Bai", "Phan", "Do kho"
        ]
        
        # Check if the file has 10 columns
        if len(df.columns) >= 10:
            # Rename the columns
            df.columns = column_names
            st.success("Tệp đã được tải lên và kiểm tra thành công!")
            st.session_state.question_library = df
            st.session_state.file_uploaded = True
            st.session_state.file_valid = True
            
            # Display a preview of the uploaded data
            st.subheader("Xem trước thư viện câu hỏi")
            st.dataframe(df.head(5))
            
            # Display total number of questions in the library
            st.info(f"Tổng số câu hỏi trong thư viện: {len(df)}")
        else:
            st.error(f"Tệp đã tải lên có {len(df.columns)} cột. Nó phải có ít nhất 10 cột.")
            st.session_state.file_valid = False
    except Exception as e:
        st.error(f"Error reading the file: {str(e)}")
        st.session_state.file_valid = False

# Question Selection and Export section (only show if file is valid)
if st.session_state.file_valid:
    st.header("Tạo đề thi")
    
    # Number of questions selector
    num_questions = st.number_input(
        "Number of questions to select",
        min_value=1,
        max_value=len(st.session_state.question_library),
        value=min(40, len(st.session_state.question_library)),
        step=1
    )
    
    # Option to underline correct answers
    underline_correct = st.checkbox("Gạch chân đáp án đúng", value=True)
    
    # Input for space between questions
    space_between_questions = st.number_input(
        "Khoảng cách giữa các câu (điểm)", 
        min_value=0, 
        max_value=30,
        value=0,
        help="Đặt khoảng cách giữa các câu hỏi theo điểm (72 điểm = 1 inch). Giá trị cao hơn sẽ tạo ra khoảng cách lớn hơn giữa các câu hỏi."
    )
    
    # Generate button
    if st.button("Tạo câu hỏi ngẫu nhiên"):
        if num_questions > len(st.session_state.question_library):
            st.warning(f"Bạn đã yêu cầu {num_questions} câu hỏi, nhưng thư viện chỉ có {len(st.session_state.question_library)} câu hỏi. Tất cả câu hỏi sẽ được chọn.")
            selected_questions = st.session_state.question_library
        else:
            # Randomly select questions
            selected_indices = np.random.choice(
                len(st.session_state.question_library),
                size=num_questions,
                replace=False
            )
            selected_questions = st.session_state.question_library.iloc[selected_indices].reset_index(drop=True)
        
        # Store the selected questions in session state
        st.session_state.selected_questions = selected_questions
        
        # Display the selected questions
        st.subheader("Selected Questions")
        st.dataframe(selected_questions)
        
        # Create Word document
        doc = Document()
        doc.add_heading('Đề thi', level=1)
        
        # Apply single line spacing with no spacing before/after
        style = doc.styles['Normal']
        style.paragraph_format.line_spacing = 1.0  # Single spacing
        style.paragraph_format.space_before = 0
        style.paragraph_format.space_after = 0
        
        # Store answer key information for adding at the end
        answer_key = []
        
        # Add each question to the document
        for i, row in enumerate(selected_questions.itertuples(), 1):
            # Create a single paragraph for question number and text
            question_para = doc.add_paragraph()
            question_para.paragraph_format.line_spacing = 1.0
            # Add space before if not the first question
            if i > 1:
                # Convert points to Pt units (20 = 1 Pt in python-docx)
                question_para.paragraph_format.space_before = Pt(space_between_questions)
            else:
                question_para.paragraph_format.space_before = 0
            question_para.paragraph_format.space_after = 0
            
            # Add the question number in bold
            question_number = question_para.add_run(f"Câu {i}: ")
            question_number.bold = True
            
            # Add the question text in italic
            question_text = question_para.add_run(row.Cau_hoi if hasattr(row, 'Cau_hoi') else row[2])  # Get question text
            question_text.italic = True
            
            # Get answer options and correct answer
            options = [
                ("A", row.Tra_loi_1 if hasattr(row, 'Tra_loi_1') else row[3]),
                ("B", row.Tra_loi_2 if hasattr(row, 'Tra_loi_2') else row[4]),
                ("C", row.Tra_loi_3 if hasattr(row, 'Tra_loi_3') else row[5]),
                ("D", row.Tra_loi_4 if hasattr(row, 'Tra_loi_4') else row[6])
            ]
            correct_answer = row.Dap_an_dung if hasattr(row, 'Dap_an_dung') else row[7]
            
            # Get answer options and the correct answer content
            original_answers = [
                row.Tra_loi_1 if hasattr(row, 'Tra_loi_1') else row[3],
                row.Tra_loi_2 if hasattr(row, 'Tra_loi_2') else row[4],
                row.Tra_loi_3 if hasattr(row, 'Tra_loi_3') else row[5],
                row.Tra_loi_4 if hasattr(row, 'Tra_loi_4') else row[6]
            ]
            
            # Determine the original correct answer index (0-3)
            correct_index = None
            if isinstance(correct_answer, str) and correct_answer.upper() in ["A", "B", "C", "D"]:
                # Convert A,B,C,D to 0,1,2,3
                correct_index = ord(correct_answer.upper()) - ord('A')
            elif isinstance(correct_answer, (int, float)):
                try:
                    # Convert 1,2,3,4 to 0,1,2,3
                    temp_index = int(correct_answer) - 1
                    if 0 <= temp_index <= 3:
                        correct_index = temp_index
                except (ValueError, TypeError):
                    pass
            
            # If we couldn't determine the correct index, default to 0
            if correct_index is None:
                correct_index = 0
                
            # Get the correct answer content
            correct_content = original_answers[correct_index]
            
            # Create a copy of answers for shuffling
            shuffled_answers = original_answers.copy()
            import random
            random.shuffle(shuffled_answers)
            
            # Find where the correct answer ended up after shuffling
            new_correct_index = shuffled_answers.index(correct_content)
            new_correct_letter = chr(65 + new_correct_index)  # 0->A, 1->B, etc.
            
            # Add to answer key
            answer_key.append(f"{i}{new_correct_letter}")
            
            # Create the options with A, B, C, D labels
            options = [
                ("A", shuffled_answers[0]),
                ("B", shuffled_answers[1]),
                ("C", shuffled_answers[2]),
                ("D", shuffled_answers[3])
            ]
            
            # Calculate the layout based on answer lengths
            answer_lengths = [len(content) for content in shuffled_answers]
            avg_length = sum(answer_lengths) / len(answer_lengths)
            max_length = max(answer_lengths)
            
            # Decide layout: 1 row (all answers in one paragraph), 2 rows (2x2), or 4 rows (1 per line)
            if max_length < 15 and avg_length < 10:  # Short answers - put all on one line
                answers_para = doc.add_paragraph()
                answers_para.paragraph_format.line_spacing = 1.0
                answers_para.paragraph_format.space_before = 0
                answers_para.paragraph_format.space_after = 0
                
                # Calculate the maximum width available for distribution
                page_width_points = 450  # Approximation of available space in points
                # Calculate the total space to distribute (after subtracting answer text)
                total_answer_text_len = sum(len(option_text) for _, option_text in options)
                total_letter_len = 8  # A. B. C. D. including spaces (2 chars each)
                
                # Space available to distribute between answers
                space_to_distribute = max(20, page_width_points - total_answer_text_len - total_letter_len)
                # Space between each answer (3 spaces)
                space_between = min(20, int(space_to_distribute / 3))  # Limit to max 20 spaces
                # Create a spacing string with the appropriate number of spaces
                spacing = " " * space_between
                
                for idx, (option_letter, option_text) in enumerate(options):
                    # Add even spacing between options
                    if idx > 0:
                        # Use the calculated spacing
                        answers_para.add_run(spacing)
                    
                    # Make the option letter bold
                    option_label = answers_para.add_run(f"{option_letter}. ")
                    option_label.bold = True
                    
                    # Check if this option matches the correct answer
                    is_correct = (option_letter == new_correct_letter)
                    
                    # Add the answer text, underlined if it's correct
                    option_content = answers_para.add_run(option_text)
                    if is_correct and underline_correct:
                        option_content.underline = True
                        
            elif max_length < 40 and avg_length < 30:  # Medium length - put in 2 rows (2 answers per line)
                # First row with A and B
                row1_para = doc.add_paragraph()
                row1_para.paragraph_format.line_spacing = 1.0
                row1_para.paragraph_format.space_before = 0
                row1_para.paragraph_format.space_after = 0
                
                # Calculate spacing for this row
                page_width_points = 450  # Approximation of available space in points
                # Calculate the total space to distribute (after subtracting answer text)
                row1_text_len = sum(len(option_text) for _, option_text in options[:2])
                letter_len = 4  # A. B. including spaces (2 chars each)
                
                # Space available to distribute between answers
                space_to_distribute = max(20, page_width_points - row1_text_len - letter_len)
                # Only one gap between A and B
                space_between = min(20, int(space_to_distribute))  # Limit to max 20 spaces
                # Create a spacing string with the appropriate number of spaces
                spacing = " " * space_between
                
                # Add first two options (A and B)
                for idx in range(2):
                    if idx > 0:
                        row1_para.add_run(spacing)
                    
                    option_letter, option_text = options[idx]
                    option_label = row1_para.add_run(f"{option_letter}. ")
                    option_label.bold = True
                    
                    is_correct = (option_letter == new_correct_letter)
                    
                    option_content = row1_para.add_run(option_text)
                    if is_correct and underline_correct:
                        option_content.underline = True
                
                # Second row with C and D
                row2_para = doc.add_paragraph()
                row2_para.paragraph_format.line_spacing = 1.0
                row2_para.paragraph_format.space_before = 0
                row2_para.paragraph_format.space_after = 0
                
                # Calculate spacing for this row
                # Calculate the total space to distribute (after subtracting answer text)
                row2_text_len = sum(len(option_text) for _, option_text in options[2:4])
                letter_len = 4  # C. D. including spaces (2 chars each)
                
                # Space available to distribute between answers
                space_to_distribute = max(20, page_width_points - row2_text_len - letter_len)
                # Only one gap between C and D
                space_between = min(20, int(space_to_distribute))  # Limit to max 20 spaces
                # Create a spacing string with the appropriate number of spaces
                spacing = " " * space_between
                
                # Add last two options (C and D)
                for idx in range(2, 4):
                    if idx > 2:
                        row2_para.add_run(spacing)
                    
                    option_letter, option_text = options[idx]
                    option_label = row2_para.add_run(f"{option_letter}. ")
                    option_label.bold = True
                    
                    is_correct = (option_letter == new_correct_letter)
                    
                    option_content = row2_para.add_run(option_text)
                    if is_correct and underline_correct:
                        option_content.underline = True
            
            else:  # Long answers - put each on its own line
                for option_letter, option_text in options:
                    answer_para = doc.add_paragraph()
                    answer_para.paragraph_format.line_spacing = 1.0
                    answer_para.paragraph_format.space_before = 0
                    answer_para.paragraph_format.space_after = 0
                    
                    # Make the option letter bold
                    option_label = answer_para.add_run(f"{option_letter}. ")
                    option_label.bold = True
                    
                    # Check if this option matches the correct answer
                    is_correct = (option_letter == new_correct_letter)
                    
                    # Add the answer text, underlined if it's correct
                    option_content = answer_para.add_run(option_text)
                    if is_correct and underline_correct:
                        option_content.underline = True
        
        # Add the answer key at the end of the document
        if answer_key:
            # Add a page break
            doc.add_page_break()
            
            # Add answer key heading
            doc.add_heading('Đáp án', level=1)
            
            # Add answer key content
            answer_key_para = doc.add_paragraph()
            answer_key_para.paragraph_format.line_spacing = 1.0
            answer_key_para.paragraph_format.space_before = 0
            answer_key_para.paragraph_format.space_after = 0
            
            # Calculate spacing for the answer key
            answers_per_line = 5
            # Approximate page width
            page_width_points = 450  
            # Each answer key entry is approx 3 chars (e.g., "1A "), and we want 5 per line
            entry_width = 3 * answers_per_line
            # Space to distribute
            space_to_distribute = max(20, page_width_points - entry_width)
            # Spaces between each answer (4 gaps for 5 answers)
            space_between = min(20, int(space_to_distribute / (answers_per_line - 1)))  # Limit to max 20 spaces
            # Create a spacing string
            spacing = " " * space_between
            
            # Format with calculated spacing
            for i, answer in enumerate(answer_key):
                if i > 0 and i % answers_per_line == 0:
                    answer_key_para.add_run('\n')
                elif i > 0:
                    answer_key_para.add_run(spacing)
                answer_key_para.add_run(answer)
        
        # Save the document to a BytesIO object
        docx_io = io.BytesIO()
        doc.save(docx_io)
        docx_io.seek(0)
        
        # Offer the document for download
        st.download_button(
            label="Tải xuống tài liệu Word",
            data=docx_io,
            file_name="examination_questions.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# Add instructions if no file has been uploaded
if not st.session_state.file_uploaded:
    st.info("Vui lòng tải lên tệp Excel chứa thư viện câu hỏi để bắt đầu.")

# Add footer with instructions
st.markdown("---")
st.markdown("""
### Hướng dẫn sử dụng:
1. Tải lên tệp Excel chứa thư viện câu hỏi (phải có ít nhất 10 cột).
2. Chọn số lượng câu hỏi bạn muốn tạo (mặc định là 40).
3. Đặt khoảng cách giữa các câu hỏi (0-30 điểm, trong đó 72 điểm = 1 inch).
4. Chọn có gạch chân đáp án đúng hay không.
5. Nhấp vào "Tạo câu hỏi ngẫu nhiên" để chọn ngẫu nhiên câu hỏi từ thư viện.
6. Xem trước các câu hỏi đã chọn.
7. Tải xuống câu hỏi dưới dạng tài liệu Word.

### Word Document Format:
- Questions formatted with "Câu [number]:" and question text on the same line
- Question number in bold and question text in italic
- Answers labeled as A, B, C, D in bold (with randomized content)
- Option to underline the correct answer (can be toggled on/off)
- Single line spacing with customizable space between questions
- Answers maximally distributed across the page while keeping them in the same row
- Answer key included at the end of the document (format: 1A, 2B, 3D, etc.)
""")

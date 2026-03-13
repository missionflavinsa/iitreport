"""
IIT Report Generation - Streamlit Application
Generate student scorecards from multiple IIT Foundation test results
"""

import streamlit as st
import pandas as pd
from io import BytesIO
import time

# Page configuration
st.set_page_config(
    page_title="IIT Report Generation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto"
)

# Import utilities
from utils.data_processor import read_excel_file, merge_all_tests, get_all_students_list
from utils.report_generator import (
    create_scorecard_excel, get_download_filename,
    create_word_report_cards, get_word_filename,
    create_consolidated_excel, get_consolidated_filename
)

# CSS with theme support
st.markdown("""
<style>
    :root {
        --bg-info: #E8F4FD;
        --text-primary: #1E3A5F;
        --text-secondary: #666666;
        --border-info: #B8DAFF;
        --btn-primary: #1E3A5F;
        --card-bg: #F8F9FA;
    }
    
    @media (prefers-color-scheme: dark) {
        :root {
            --bg-info: #1E2A3A;
            --text-primary: #FAFAFA;
            --text-secondary: #B0B0B0;
            --border-info: #3A5A7A;
            --btn-primary: #4A90D9;
            --card-bg: #262730;
        }
    }
    
    .main-header {
        font-size: clamp(1.5rem, 5vw, 2.5rem);
        font-weight: 700;
        color: var(--text-primary);
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: clamp(0.9rem, 2.5vw, 1.1rem);
        color: var(--text-secondary);
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .info-box {
        background-color: var(--bg-info);
        border: 1px solid var(--border-info);
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
        color: var(--text-primary);
    }
    .metric-card {
        background-color: var(--card-bg);
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        margin-bottom: 1rem;
    }
    .block-container { padding-top: 2rem !important; }
    .footer { text-align: center; color: #888; font-size: 0.85rem; padding: 1rem; }
    
    @media (max-width: 768px) {
        [data-testid="column"] { width: 100% !important; flex: 1 1 100% !important; }
        .main .block-container { padding: 1rem 0.5rem !important; }
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<h1 class="main-header">📊 IIT Report Generation</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Generate student scorecards from IIT Foundation test results</p>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("📋 Configuration")
    class_name = st.text_input("Class", placeholder="e.g., VII, IX")
    section = st.text_input("Section", placeholder="e.g., A, B, KC")
    st.divider()
    st.subheader("📁 Upload Test Files")
    uploaded_files = st.file_uploader(
        "Upload Excel Files",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Upload IIT test result files (.xlsx)"
    )

# Main content
if not uploaded_files:
    st.info("👈 Upload IIT test Excel files from the sidebar to begin.")
    
    cols = st.columns(3)
    steps = [("📤", "Step 1", "Enter class & section"),
             ("📁", "Step 2", "Upload test files"),
             ("📥", "Step 3", "Generate scorecard")]
    for col, (icon, title, desc) in zip(cols, steps):
        with col:
            st.markdown(f'<div class="metric-card"><h3>{icon} {title}</h3><p>{desc}</p></div>', unsafe_allow_html=True)
    
    with st.expander("📋 Expected Excel Format"):
        st.markdown("""
        **Headers:** `Sr. No.`, `Candidate ID`, `Name of the Student`, `Phy`, `Chem`, `Maths`, `Bio`, `Total`
        
        The app auto-detects the header row in your Excel files.
        """)

else:
    st.success(f"✅ {len(uploaded_files)} file(s) selected for processing")
    
    # Initialize containers
    progress_container = st.container()
    results_container = st.container()
    
    with progress_container:
        st.subheader("📊 Processing Files")
        
        # Overall progress
        overall_progress = st.progress(0, text="Starting...")
        
        # Status expander
        status_expander = st.expander("📋 Processing Details", expanded=True)
        
        files_data = []
        file_info = []
        errors = []
        
        total_files = len(uploaded_files)
        
        with status_expander:
            for i, file in enumerate(uploaded_files):
                file_progress = (i / total_files)
                overall_progress.progress(file_progress, text=f"Processing file {i+1}/{total_files}: {file.name[:40]}...")
                
                # Show current file being processed
                status_placeholder = st.empty()
                status_placeholder.info(f"🔄 Reading: **{file.name}**")
                
                try:
                    # Read file content
                    file_content = file.read()
                    
                    # Process file
                    df, test_date = read_excel_file(file_content, file.name)
                    
                    files_data.append((df, test_date))
                    
                    file_info.append({
                        'File': file.name[:40] + "..." if len(file.name) > 40 else file.name,
                        'Date': test_date,
                        'Students': len(df),
                        'Status': '✅ Success'
                    })
                    
                    status_placeholder.success(f"✅ **{file.name}** - {len(df)} students found (Date: {test_date})")
                    
                except Exception as e:
                    errors.append(f"{file.name}: {str(e)}")
                    file_info.append({
                        'File': file.name[:40] + "..." if len(file.name) > 40 else file.name,
                        'Date': 'Error',
                        'Students': 0,
                        'Status': '❌ Failed'
                    })
                    status_placeholder.error(f"❌ **{file.name}** - Error: {str(e)}")
                
                time.sleep(0.1)  # Brief pause for visual feedback
        
        overall_progress.progress(1.0, text="✅ Processing complete!")
    
    # Show results
    with results_container:
        if errors:
            st.warning(f"⚠️ {len(errors)} file(s) had errors. Check details above.")
        
        st.divider()
        
        # Files summary table
        st.subheader("📋 Files Summary")
        summary_df = pd.DataFrame(file_info)
        st.dataframe(summary_df, hide_index=True, width='stretch')
        
        # Process successfully loaded files
        if files_data:
            st.divider()
            
            # Merge data progress
            merge_progress = st.progress(0, text="Merging test data...")
            all_students, test_dates = merge_all_tests(files_data)
            merge_progress.progress(1.0, text="✅ Data merged successfully!")
            
            # Statistics
            col1, col2, col3 = st.columns(3)
            col1.metric("📚 Total Tests", len(test_dates))
            col2.metric("👨‍🎓 Total Students", len(all_students))
            col3.metric("📅 Date Range", f"{test_dates[0]} - {test_dates[-1]}" if len(test_dates) > 1 else (test_dates[0] if test_dates else "N/A"))
            
            st.divider()
            
            # Student list
            students_list = get_all_students_list(all_students)
            
            with st.expander(f"👨‍🎓 View {len(students_list)} Students"):
                students_df = pd.DataFrame(students_list)
                students_df.columns = ['Candidate ID', 'Student Name']
                st.dataframe(students_df, hide_index=True, height=300, width='stretch')
            
            st.divider()
            
            # Generate section
            st.subheader("📊 Generate Scorecard")
            
            if not class_name or not section:
                st.warning("⚠️ Please enter **Class** and **Section** in the sidebar to generate reports.")
            else:
                st.markdown(f"""
                <div class="info-box">
                    <strong>Ready to generate reports for:</strong> Class <strong>{class_name}</strong> - Section <strong>{section}</strong><br>
                    <small>Choose the report type below to download.</small>
                </div>
                """, unsafe_allow_html=True)
                
                # Academic year input
                academic_year = st.text_input("Academic Year", value="2024-25", help="e.g., 2024-25")
                
                st.divider()
                
                # Three columns for download buttons
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("**📊 Individual Scorecards**")
                    st.caption("Excel with one sheet per student + charts")
                    if st.button("Generate Scorecards", key="btn_scorecard", use_container_width=True):
                        with st.spinner("Generating..."):
                            try:
                                excel_buffer = create_scorecard_excel(all_students, test_dates, class_name, section)
                                filename = get_download_filename(class_name, section)
                                st.download_button(
                                    label="📥 Download Excel",
                                    data=excel_buffer,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_scorecard"
                                )
                                st.success("✅ Ready!")
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                
                with col2:
                    st.markdown("**📄 Report Cards (Word)**")
                    st.caption("One page per student, print-ready")
                    if st.button("Generate Word Doc", key="btn_word", use_container_width=True):
                        with st.spinner("Generating..."):
                            try:
                                word_buffer = create_word_report_cards(all_students, test_dates, class_name, section, academic_year)
                                filename = get_word_filename(class_name, section)
                                st.download_button(
                                    label="📥 Download Word",
                                    data=word_buffer,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    key="dl_word"
                                )
                                st.success("✅ Ready!")
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
                
                with col3:
                    st.markdown("**📋 Consolidated Report**")
                    st.caption("All students, all tests in one sheet")
                    if st.button("Generate Consolidated", key="btn_consolidated", use_container_width=True):
                        with st.spinner("Generating..."):
                            try:
                                excel_buffer = create_consolidated_excel(all_students, test_dates, class_name, section, academic_year)
                                filename = get_consolidated_filename(class_name, section)
                                st.download_button(
                                    label="📥 Download Excel",
                                    data=excel_buffer,
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="dl_consolidated"
                                )
                                st.success("✅ Ready!")
                            except Exception as e:
                                st.error(f"Error: {str(e)}")
        else:
            st.error("❌ No files were processed successfully. Please check your Excel files.")

# Footer
st.markdown('<div class="footer">IIT Report Generation | Rotary English Medium School</div>', unsafe_allow_html=True)

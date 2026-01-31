import streamlit as st
import pandas as pd
import os
import asyncio
import edge_tts
import fitz  # PyMuPDF
import comtypes.client
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
import tempfile
import shutil

# --- PAGE CONFIGURATION (Must be first) ---
st.set_page_config(
    page_title="L&D Video Studio",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM CSS FOR AESTHETICS ---
st.markdown("""
<style>
    /* Main container padding */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    /* Card-like borders for containers */
    [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
        border: 1px solid #f0f2f6;
        border-radius: 10px;
        padding: 1rem;
        background-color: white;
    }
    /* Custom Button Styling */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
    /* specific style for the primary action button */
    div.stButton > button:first-child {
        background-color: #FF4B4B;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# --- CONFIGURATION & CONSTANTS ---
TEMP_DIR = "temp_processing"
OUTPUT_DIR = "generated_videos"

VOICE_OPTIONS = {
    "🇮🇳 Neerja (Indian Female)": "en-IN-NeerjaNeural",
    "🇮🇳 Prabhat (Indian Male)": "en-IN-PrabhatNeural",
    "🇺🇸 Aria (US Female)": "en-US-AriaNeural",
    "🇺🇸 Christopher (US Male)": "en-US-ChristopherNeural",
    "🇬🇧 Sonia (UK Female)": "en-GB-SoniaNeural",
    "🇬🇧 Ryan (UK Male)": "en-GB-RyanNeural"
}

SPEED_OPTIONS = {
    "Slow 🐢": "-20%",
    "Normal 🐇": "+0%",
    "Fast 🐆": "+20%"
}

# Ensure directories exist
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- HELPER FUNCTIONS ---

def parse_excel_script(uploaded_excel):
    try:
        df = pd.read_excel(uploaded_excel)
        if 'Script' not in df.columns:
            st.error("❌ Excel must have a column named 'Script'.")
            return []
        return df['Script'].dropna().astype(str).tolist()
    except Exception as e:
        st.error(f"❌ Error reading Excel: {e}")
        return []

def save_uploaded_file(uploaded_file):
    file_path = os.path.join(TEMP_DIR, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path

def convert_pdf_to_images(pdf_path, output_folder):
    doc = fitz.open(pdf_path)
    zoom_matrix = fitz.Matrix(2.0, 2.0)
    image_paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=zoom_matrix)
        img_name = f"slide_{i:03d}.png"
        img_path = os.path.join(output_folder, img_name)
        pix.save(img_path)
        image_paths.append(img_path)
    return image_paths

def convert_ppt_to_images_windows(ppt_path, output_folder):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    # powerpoint.Visible = 1 # Uncomment if needed for debugging
    abs_ppt_path = os.path.abspath(ppt_path)
    abs_output_folder = os.path.abspath(output_folder)
    try:
        presentation = powerpoint.Presentations.Open(abs_ppt_path)
        image_paths = []
        for i, slide in enumerate(presentation.Slides):
            img_name = f"slide_{i:03d}.png"
            img_path = os.path.join(abs_output_folder, img_name)
            slide.Export(img_path, "PNG", 1920, 1080)
            image_paths.append(img_path)
        presentation.Close()
        return image_paths
    except Exception as e:
        st.error(f"PPT Conversion Error: {e}")
        return []
    finally:
        try:
            powerpoint.Quit()
        except:
            pass

async def generate_single_audio(text, voice, speed, output_path):
    communicate = edge_tts.Communicate(text, voice, rate=speed)
    await communicate.save(output_path)

def generate_video_pipeline(voice_name, voice_id, speed_rate, script_lines, image_paths, status_placeholder, progress_callback):
    # Clean inputs for filename safety
    safe_voice_name = voice_name.split()[1]
    safe_speed = speed_rate.replace('%', '').replace('+', '')
    
    run_id = f"{safe_voice_name}_{safe_speed}"
    voice_folder = os.path.join(TEMP_DIR, run_id)
    os.makedirs(voice_folder, exist_ok=True)

    audio_paths = []
    
    # --- STEP 1: AUDIO GENERATION ---
    status_placeholder.write(f"🎙️ **{voice_name}**: Synthesizing audio...")
    
    async def run_audio_gen():
        tasks = []
        for idx, text in enumerate(script_lines):
            if idx >= len(image_paths): break
            audio_path = os.path.join(voice_folder, f"audio_{idx:03d}.mp3")
            audio_paths.append(audio_path)
            
            # Generate Audio
            await generate_single_audio(text, voice_id, speed_rate, audio_path)
            
            # Update Progress
            progress_callback()
            
    asyncio.run(run_audio_gen())

    # --- STEP 2: VIDEO RENDERING ---
    status_placeholder.write(f"🎬 **{voice_name}**: Rendering video timeline...")
    
    clips = []
    for img, audio in zip(image_paths, audio_paths):
        try:
            audio_clip = AudioFileClip(audio)
            video_clip = ImageClip(img).set_duration(audio_clip.duration).set_audio(audio_clip)
            clips.append(video_clip)
        except Exception as e:
            st.warning(f"Skipping slide due to error: {e}")

    if not clips: return None

    final_filename = f"L&D_Video_{safe_voice_name}_{safe_speed}.mp4"
    final_path = os.path.join(OUTPUT_DIR, final_filename)
    
    try:
        final_video = concatenate_videoclips(clips, method="compose")
        final_video.write_videofile(final_path, fps=24, codec="libx264", audio_codec="aac", logger=None)
        return final_path
    except Exception as e:
        st.error(f"Render Error: {e}")
        return None

# --- MAIN UI LOGIC ---

# Header Section
st.markdown("<h1 style='text-align: center; color: #333;'>🎬 L&D Automation Studio</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666;'>Turn slides and scripts into professional training videos in minutes.</p>", unsafe_allow_html=True)
st.divider()

if 'generated_videos' not in st.session_state:
    st.session_state['generated_videos'] = []

# Layout: Left for Config, Right for Action/Results
left_col, right_col = st.columns([1, 1.5], gap="large")

with left_col:
    # --- SECTION 1: UPLOAD ---
    with st.container(border=True):
        st.subheader("📂 1. Project Files")
        uploaded_slides = st.file_uploader("Upload Presentation (PDF/PPTX)", type=["pdf", "pptx"])
        uploaded_excel = st.file_uploader("Upload Script (Excel)", type=["xlsx"])
        
        with st.expander("ℹ️ View File Format Instructions"):
            st.markdown("""
            **Excel Format:**
            - **Column A:** `Slides` (1, 2, 3...)
            - **Column B:** `Script` (The text to be spoken)
            
            **Presentation:**
            - Ensure slide count matches script rows.
            """)

    # --- SECTION 2: SETTINGS ---
    st.write("") # Spacer
    with st.container(border=True):
        st.subheader("🎛️ 2. Voice Studio")
        
        selected_voices = st.multiselect(
            "Select Narrators (Max 4)", 
            options=list(VOICE_OPTIONS.keys()), 
            default=["🇮🇳 Prabhat (Indian Male)"],
            max_selections=4
        )
        
        selected_speed_label = st.select_slider(
            "Narrator Speed", 
            options=list(SPEED_OPTIONS.keys()), 
            value="Normal 🐇"
        )
        selected_speed_val = SPEED_OPTIONS[selected_speed_label]

        # --- PREVIEW LAB ---
        with st.expander("🎧 Audio Lab (Test Voices)"):
            preview_voice = st.selectbox("Select Voice to Test", list(VOICE_OPTIONS.keys()))
            preview_text = st.text_input("Test Text", "Welcome to the L&D automation module.")
            
            if st.button("▶️ Play Preview"):
                with st.spinner("Synthesizing..."):
                    preview_file = "preview.mp3"
                    v_id = VOICE_OPTIONS[preview_voice]
                    asyncio.run(generate_single_audio(preview_text, v_id, selected_speed_val, preview_file))
                    st.audio(preview_file)

with right_col:
    # --- SECTION 3: ACTION DASHBOARD ---
    with st.container(border=True):
        st.subheader("🚀 3. Production Dashboard")
        
        start_btn = st.button("Generate Videos", type="primary", use_container_width=True)
        
        # Dashboard Area
        status_area = st.empty()
        progress_bar = st.progress(0)
        
        if start_btn:
            # RESET
            st.session_state['generated_videos'] = []
            progress_bar.progress(0)
            
            # VALIDATION
            if not uploaded_slides or not uploaded_excel:
                status_area.error("⚠️ Please upload both Presentation and Excel Script files first.")
            elif not selected_voices:
                status_area.error("⚠️ Please select at least one narrator.")
            else:
                # EXECUTION
                with st.status("🏗️ Production in progress...", expanded=True) as status:
                    
                    # A. Parse Script
                    status.write("📄 Reading Script...")
                    script_lines = parse_excel_script(uploaded_excel)
                    
                    if script_lines:
                        # B. Process Slides
                        status.write("🖼️ Processing Slides...")
                        slides_temp_dir = os.path.join(TEMP_DIR, "slides_images")
                        os.makedirs(slides_temp_dir, exist_ok=True)
                        slide_path = save_uploaded_file(uploaded_slides)
                        
                        if slide_path.endswith(".pdf"):
                            image_paths = convert_pdf_to_images(slide_path, slides_temp_dir)
                        else:
                            image_paths = convert_ppt_to_images_windows(slide_path, slides_temp_dir)
                        
                        # C. Pipeline Setup
                        total_steps = len(script_lines) * len(selected_voices)
                        step_counter = [0] # Mutable list for progress tracking

                        def update_progress():
                            step_counter[0] += 1
                            pct = min(step_counter[0] / total_steps, 1.0)
                            progress_bar.progress(pct)

                        # D. Generation Loop
                        for i, v_name in enumerate(selected_voices):
                            v_id = VOICE_OPTIONS[v_name]
                            
                            video_path = generate_video_pipeline(
                                v_name, v_id, selected_speed_val, 
                                script_lines, image_paths, status, update_progress
                            )
                            
                            if video_path:
                                st.session_state['generated_videos'].append((v_name, video_path))

                        status.update(label="✅ Production Complete!", state="complete", expanded=False)
                        status_area.success("🎉 All videos generated successfully!")

    # --- SECTION 4: DOWNLOADS ---
    if st.session_state['generated_videos']:
        st.write("")
        st.subheader("📥 Ready for Download")
        
        # Display in a grid
        cols = st.columns(2)
        for idx, (v_name, file_path) in enumerate(st.session_state['generated_videos']):
            if os.path.exists(file_path):
                with open(file_path, "rb") as file:
                    # Alternating columns for grid effect
                    with cols[idx % 2]:
                        st.download_button(
                            label=f"⬇️ {v_name}",
                            data=file,
                            file_name=os.path.basename(file_path),
                            mime="video/mp4",
                            key=f"dl_btn_{idx}",
                            use_container_width=True
                        )
            
        st.divider()
        if st.button("🔄 Start New Project"):
            st.session_state['generated_videos'] = []
            st.rerun()

# Footer
st.markdown("<div style='text-align: center; margin-top: 50px; color: #aaa; font-size: 0.8em;'>L&D Automation Tool v1.2</div>", unsafe_allow_html=True)
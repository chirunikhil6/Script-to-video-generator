import streamlit as st
import pandas as pd
import os
import asyncio
import edge_tts
import fitz  # PyMuPDF
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
import tempfile
import shutil

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="L&D Video Studio",
    page_icon="🎬",
    layout="wide"
)

# --- CSS ---
st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    div.stButton > button:first-child { background-color: #FF4B4B; color: white; }
</style>
""", unsafe_allow_html=True)

# --- CONFIG ---
TEMP_DIR = "temp_processing"
OUTPUT_DIR = "generated_videos"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

VOICE_OPTIONS = {
    "🇮🇳 Neerja (Indian Female)": "en-IN-NeerjaNeural",
    "🇮🇳 Prabhat (Indian Male)": "en-IN-PrabhatNeural",
    "🇺🇸 Aria (US Female)": "en-US-AriaNeural",
    "🇺🇸 Christopher (US Male)": "en-US-ChristopherNeural",
    "🇬🇧 Sonia (UK Female)": "en-GB-SoniaNeural",
    "🇬🇧 Ryan (UK Male)": "en-GB-RyanNeural"
}

SPEED_OPTIONS = { "Slow 🐢": "-20%", "Normal 🐇": "+0%", "Fast 🐆": "+20%" }

# --- FUNCTIONS ---

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

async def generate_single_audio(text, voice, speed, output_path):
    communicate = edge_tts.Communicate(text, voice, rate=speed)
    await communicate.save(output_path)

def generate_video_pipeline(voice_name, voice_id, speed_rate, script_lines, image_paths, status_placeholder, progress_callback):
    safe_voice_name = voice_name.split()[1]
    safe_speed = speed_rate.replace('%', '').replace('+', '')
    run_id = f"{safe_voice_name}_{safe_speed}"
    voice_folder = os.path.join(TEMP_DIR, run_id)
    os.makedirs(voice_folder, exist_ok=True)

    audio_paths = []
    status_placeholder.write(f"🎙️ **{voice_name}**: Synthesizing audio...")
    
    async def run_audio_gen():
        tasks = []
        for idx, text in enumerate(script_lines):
            if idx >= len(image_paths): break
            audio_path = os.path.join(voice_folder, f"audio_{idx:03d}.mp3")
            audio_paths.append(audio_path)
            await generate_single_audio(text, voice_id, speed_rate, audio_path)
            progress_callback()
    asyncio.run(run_audio_gen())

    status_placeholder.write(f"🎬 **{voice_name}**: Rendering video...")
    clips = []
    for img, audio in zip(image_paths, audio_paths):
        try:
            audio_clip = AudioFileClip(audio)
            video_clip = ImageClip(img).set_duration(audio_clip.duration).set_audio(audio_clip)
            clips.append(video_clip)
        except Exception as e:
            st.warning(f"Skipping slide error: {e}")

    if not clips: return None

    final_filename = f"Video_{safe_voice_name}.mp4"
    final_path = os.path.join(OUTPUT_DIR, final_filename)
    final_video = concatenate_videoclips(clips, method="compose")
    final_video.write_videofile(final_path, fps=24, codec="libx264", audio_codec="aac", logger=None)
    return final_path

# --- UI ---
st.title("🎬 L&D Automation Studio (Cloud)")
st.markdown("Upload **PDF Slides** and **Excel Script** to generate videos.")

if 'generated_videos' not in st.session_state:
    st.session_state['generated_videos'] = []

left_col, right_col = st.columns([1, 1.5], gap="large")

with left_col:
    with st.container(border=True):
        st.subheader("📂 1. Upload Files")
        # NOTE: Restricted to PDF only for Cloud compatibility
        uploaded_slides = st.file_uploader("Upload Slides (PDF Only)", type=["pdf"])
        uploaded_excel = st.file_uploader("Upload Script (Excel)", type=["xlsx"])

    with st.container(border=True):
        st.subheader("🎛️ 2. Settings")
        selected_voices = st.multiselect("Select Narrators", list(VOICE_OPTIONS.keys()), default=["🇮🇳 Prabhat (Indian Male)"])
        selected_speed = st.select_slider("Speed", list(SPEED_OPTIONS.keys()), value="Normal 🐇")

with right_col:
    with st.container(border=True):
        st.subheader("🚀 3. Produce")
        if st.button("Generate Videos", type="primary", use_container_width=True):
            st.session_state['generated_videos'] = []
            if not uploaded_slides or not uploaded_excel:
                st.error("Upload both files first.")
            else:
                with st.status("Processing...", expanded=True) as status:
                    script_lines = parse_excel_script(uploaded_excel)
                    if script_lines:
                        slides_dir = os.path.join(TEMP_DIR, "slides")
                        save_uploaded_file(uploaded_slides) # Save pdf
                        image_paths = convert_pdf_to_images(os.path.join(TEMP_DIR, uploaded_slides.name), slides_dir)
                        
                        total_steps = len(script_lines) * len(selected_voices)
                        step_counter = [0]
                        def update_progress():
                            step_counter[0] += 1

                        for v_name in selected_voices:
                            v_id = VOICE_OPTIONS[v_name]
                            rate = SPEED_OPTIONS[selected_speed]
                            path = generate_video_pipeline(v_name, v_id, rate, script_lines, image_paths, status, update_progress)
                            if path: st.session_state['generated_videos'].append((v_name, path))
                        status.update(label="Done!", state="complete")

    if st.session_state['generated_videos']:
        st.subheader("📥 Downloads")
        for v_name, path in st.session_state['generated_videos']:
            with open(path, "rb") as f:
                st.download_button(f"⬇️ {v_name}", f, file_name=os.path.basename(path))
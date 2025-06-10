import whisper

model = whisper.load_model(
    "medium",
    download_root="C:/Users/itgee/models/whisper"  # change to your preferred directory
)

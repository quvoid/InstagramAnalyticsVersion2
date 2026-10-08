# -*- coding: utf-8 -*-
"""High-quality whisper worker: one mp3 path per stdin line -> one JSON line.
Separate process so the parent can kill a wedged decode. Model name is argv[1]."""
import json, sys
from faster_whisper import WhisperModel

MODEL = sys.argv[1] if len(sys.argv) > 1 else "large-v3-turbo"
# Brand names as a hint, so the decoder spells them the way the brands do.
HINT = ("Tanishq, Malabar Gold and Diamonds, Joyalukkas, Senco Gold, Jos Alukkas, BlueStone, "
        "Kalyan Jewellers, Thangamayil, PC Jeweller, CaratLane.")

model = WhisperModel(MODEL, device="cpu", compute_type="int8")
print(json.dumps({"ready": True, "model": MODEL}), flush=True)

for line in sys.stdin:
    path = line.strip()
    if not path:
        continue
    try:
        seg, info = model.transcribe(
            path, beam_size=5, best_of=5, vad_filter=True,
            condition_on_previous_text=False,          # stops repetition loops on music
            temperature=[0.0, 0.2, 0.4],
            compression_ratio_threshold=2.4, no_speech_threshold=0.6,
            initial_prompt=HINT,
        )
        segs = list(seg)
        txt = " ".join(s.text.strip() for s in segs).strip()
        print(json.dumps({"text": txt, "lang": info.language,
                          "lang_prob": round(info.language_probability, 2),
                          "dur": round(info.duration, 1)}, ensure_ascii=False), flush=True)
    except Exception as e:
        print(json.dumps({"error": f"{type(e).__name__}: {str(e)[:120]}"}), flush=True)

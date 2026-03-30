from google import genai

KUNCI_SAYA = "AIzaSyC4WeIpGMWa8jKxxKy-tmw84XNzR8870qM" 

print(f"Mencoba menghubungi Google Gemini Baru dengan kunci: {KUNCI_SAYA[:10]}...")

try:
    # Cara baru memanggil Google Gemini
    client = genai.Client(api_key=KUNCI_SAYA)
    response = client.models.generate_content(
        model='gemini-2.0-flash',
        contents='Halo, apakah kamu bisa mendengar saya? Jawab dalam 1 kalimat saja.'
    )
    print("\n✅ BERHASIL! Jawaban AI:")
    print(response.text)
except Exception as e:
    print("\n❌ GAGAL! Error dari Google:")
    print(e)
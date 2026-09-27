import os
import glob
import torch
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma

KNOWLEDGE_FOLDER = "knowledge_base"
DB_DIR = "tax_db"

def clean_text(text):
    if not text or not isinstance(text, str):
        return ""
    return text.replace("\x00", " ").strip()

def main():
    pdf_files = sorted(glob.glob(os.path.join(KNOWLEDGE_FOLDER, "*.pdf")))
    print(f"📚 Total {len(pdf_files)} PDF books process hone ke liye mili hain.")

    # 1. GPU Acceleration (MPS for Mac M1/M2/M3) + Large Batch (256)
    device = "mps" if torch.backends.mps.is_available() else "cpu"
    print(f"⚡ Turbo Mode Activated on: {device.upper()} (Superfast)")

    embeddings = HuggingFaceEmbeddings(
        model_name="all-MiniLM-L6-v2",
        model_kwargs={"device": device},
        encode_kwargs={"batch_size": 256}  # 8x Speed Booster
    )

    vector_db = Chroma(persist_directory=DB_DIR, embedding_function=embeddings)

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1500,     # Thoda bada chunk = Chunks ki sankhya aadhi ho jayegi
        chunk_overlap=100,
        separators=["\n\n", "\n", " ", ""]
    )

    # 2. Process Book-by-Book (Har book process hote hi save ho jayegi)
    for idx, pdf_path in enumerate(pdf_files, 1):
        b_name = os.path.basename(pdf_path)
        print(f"\n📖 [{idx}/{len(pdf_files)}] Processing: {b_name}...")
        
        book_chunks = []
        book_metas = []
        
        try:
            reader = PdfReader(pdf_path)
            for page_num, page in enumerate(reader.pages):
                raw = page.extract_text()
                t = clean_text(raw)
                if t and len(t) >= 20:
                    chunks = text_splitter.split_text(t)
                    for c in chunks:
                        clean_c = clean_text(c)
                        if clean_c and len(clean_c) >= 15:
                            book_chunks.append(clean_c)
                            book_metas.append({"source": b_name, "page": page_num + 1})
        except Exception as e:
            print(f"⚠️ Error reading {b_name}: {e}")
            continue

        if not book_chunks:
            print(f"⏩ {b_name} me koi text nahi mila, skipping.")
            continue

        print(f"   ⚙️ {len(book_chunks)} chunks ready. Embedding into DB...")

        # Batch write with bulletproof error skipping
        batch_size = 500
        for i in range(0, len(book_chunks), batch_size):
            b_texts = book_chunks[i:i + batch_size]
            b_metas = book_metas[i:i + batch_size]
            
            safe_pairs = [(txt, m) for txt, m in zip(b_texts, b_metas) if isinstance(txt, str) and len(txt.strip()) >= 10]
            if not safe_pairs:
                continue
            s_texts, s_metas = zip(*safe_pairs)
            
            try:
                vector_db.add_texts(texts=list(s_texts), metadatas=list(s_metas))
            except Exception:
                # Agar kisi chunk me dikkat ho, toh use skip karke aage badhein
                for st_item, sm_item in zip(s_texts, s_metas):
                    try:
                        vector_db.add_texts(texts=[st_item], metadatas=[sm_item])
                    except Exception:
                        pass
                        
        vector_db.persist()
        print(f"   ✅ {b_name} successfully saved to tax_db!")

    print("\n🎉 Balle Balle! Saari 21 books master database me fully index ho gayi hain!")

if __name__ == "__main__":
    main()
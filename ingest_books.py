import os
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma

# 1. Jo books add karni hain unke naam
BOOKS_TO_ADD = [
    "book1.pdf",  # Apni pehli book ka sahi naam likhein
    "book2.pdf"   # Apni doosri book ka sahi naam likhein
]

DB_DIR = "tax_db"

def extract_text_from_pdf(pdf_path):
    print(f"📖 Reading: {pdf_path}...")
    reader = PdfReader(pdf_path)
    documents = []
    for page_num, page in enumerate(reader.pages):
        text = page.extract_text()
        if text and text.strip():
            documents.append({
                "text": text,
                "metadata": {"source": os.path.basename(pdf_path), "page": page_num + 1}
            })
    return documents

def main():
    all_chunks = []
    all_metadatas = []
    
    # 2. Text Splitter (1000 characters ke tukde aur 150 overlap)
    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=150,
        separators=["\n\n", "\n", " ", ""]
    )

    for book in BOOKS_TO_ADD:
        if not os.path.exists(book):
            print(f"❌ File not found: {book}. Please check the filename!")
            continue
            
        pages = extract_text_from_pdf(book)
        for p in pages:
            chunks = text_splitter.split_text(p["text"])
            for c in chunks:
                all_chunks.append(c)
                all_metadatas.append(p["metadata"])

    if not all_chunks:
        print("⚠️ Koi data extract nahi hua. Check your PDF files.")
        return

    print(f"⚙️ Total {len(all_chunks)} chunks banaye gaye hain. Vectors banaye ja rahe hain...")

    # 3. Same Embeddings Model jo app.py me hai
    embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")

    # 4. Chroma DB me data append / add karein
    vector_db = Chroma(persist_directory=DB_DIR, embedding_function=embeddings)
    
    # Batch insertion (500 chunks at a time taaki memory crash na ho)
    batch_size = 500
    for i in range(0, len(all_chunks), batch_size):
        b_texts = all_chunks[i:i + batch_size]
        b_metas = all_metadatas[i:i + batch_size]
        vector_db.add_texts(texts=b_texts, metadatas=b_metas)
        print(f"✅ Processed {min(i + batch_size, len(all_chunks))}/{len(all_chunks)} chunks...")

    vector_db.persist()
    print("🎉 Mubarak ho! Dono books 'tax_db' me safalta-purvak add ho gayi hain.")

if __name__ == "__main__":
    main()
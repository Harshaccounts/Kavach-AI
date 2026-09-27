import os
from pypdf import PdfReader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma

folder_path = "knowledge_base"
if not os.path.exists(folder_path):
    os.makedirs(folder_path)
    print("knowledge_base folder empty hai. Kripya PDFs dalein.")
    exit()

pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
print(f"Total PDF Files Found: {len(pdf_files)}")

if len(pdf_files) == 0:
    print("Koi PDF file nahi mili! Kripya knowledge_base folder me PDF dalein.")
    exit()

embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
vector_db = Chroma(persist_directory="tax_db", embedding_function=embeddings)
text_splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=150)

for idx, file_name in enumerate(pdf_files, 1):
    file_path = os.path.join(folder_path, file_name)
    print(f"\n[{idx}/{len(pdf_files)}] Reading: {file_name}...")
    try:
        reader = PdfReader(file_path)
        total_pages = len(reader.pages)
        print(f"Total pages: {total_pages}")
        
        file_chunks = []
        for page_num, page in enumerate(reader.pages, 1):
            text = page.extract_text()
            if text:
                chunks = text_splitter.split_text(text)
                file_chunks.extend(chunks)
            if page_num % 500 == 0:
                print(f"  --> Processed {page_num}/{total_pages} pages...")
        
        print(f"Database me {len(file_chunks)} sections save ho rahe hain...")
        batch_size = 500
        for i in range(0, len(file_chunks), batch_size):
            vector_db.add_texts(file_chunks[i:i + batch_size])
        print(f"Done: {file_name}")

    except Exception as e:
        print(f"Error reading {file_name}: {e}")

print("\nSUCCESS! Aapka database 'tax_db' folder me permanently ready ho chuka hai.")
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


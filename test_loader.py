from docx import Document
import random

def load_tests_from_docx(file_path, num_questions=30):
    """
    Word (.docx) fayldan testlarni o‘qib, random 30 tasini qaytaradi.
    Har bir testda: savol, variantlar (A-D), to‘g‘ri javob.
    """
    doc = Document(file_path)
    table = doc.tables[0]
    
    tests = []
    
    for row in table.rows[1:]:  # Birinchi qator — sarlavha, o'tkazib yuboramiz
        cells = row.cells
        try:
            question_text = cells[0].text.strip()
            correct_answer = cells[1].text.strip()
            wrong_answers = [cells[2].text.strip(), cells[3].text.strip(), cells[4].text.strip()]
            
            # Variantlarni aralashtiramiz
            all_answers = wrong_answers + [correct_answer]
            random.shuffle(all_answers)
            
            tests.append({
                "question": question_text,
                "options": all_answers,
                "correct": correct_answer
            })
        except IndexError:
            print("⚠️ Xatolik: biror qatorda ustun yetarli emas.")
            continue
    
    return random.sample(tests, min(num_questions, len(tests)))


# 🧪 SINOV QISMI
if __name__ == "__main__":
    FILE_PATH = "C:/Users/FAZLIDDIN/Desktop/TEST BAZA/test_baza.docx"
    
    print("📥 Testlar yuklanmoqda...\n")
    tests = load_tests_from_docx(FILE_PATH)

    for i, t in enumerate(tests, 1):
        print(f"{i}. {t['question']}")
        for j, opt in enumerate(t['options']):
            print(f"   {chr(65 + j)}) {opt}")
        print(f"✅ To‘g‘ri javob: {t['correct']}")
        print("===")

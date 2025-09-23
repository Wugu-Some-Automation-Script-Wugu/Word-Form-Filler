import re
from typing import List, Tuple
from docx import Document
import os

def parse_source_document_test(doc_path: str) -> List[Tuple[str, str, str]]:
    try:
        doc = Document(doc_path)
        questions = []
        
        print("正在從 Word 文檔中解析內容...")
        
        full_text = ""
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text:
                full_text += text + "\n"
        
        print(f"文檔內容長度: {len(full_text)} 字元")
        print("文檔內容預覽:")
        print("-" * 50)
        print(full_text[:500] + "..." if len(full_text) > 500 else full_text)
        print("-" * 50)
        
        questions = parse_questions_from_text(full_text)
        
        print(f"成功解析 {len(questions)} 個題目")
        return questions
        
    except Exception as e:
        print(f"解析文檔時發生錯誤: {str(e)}")
        return []

def parse_questions_from_text(text: str) -> List[Tuple[str, str, str]]:
    questions = []
    lines = text.split('\n')
    
    current_question_num = 1
    current_answer = None
    current_explanation = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 識別答案行：支援單選和多選格式
        # 單選：答案 : (A) 或 答案 : (Ｃ)
        # 多選：答案 : (1)(A); (2)(D) 或 答案 : (1)(Ｂ); (2)(Ａ)
        if re.match(r'^答案\s*[：:]\s*[（(]', line):
            if current_answer:
                questions.append((
                    f"{current_question_num}.",
                    current_answer,
                    current_explanation or ""
                ))
                print(f"解析題目 {current_question_num}.: 答案={current_answer}, 解析={'有' if current_explanation else '無'}")
                current_question_num += 1
            
            # 開始新題目
            current_answer = line
            current_explanation = None 
        elif re.match(r'^解析：', line) and current_answer:
            current_explanation = line
        elif current_explanation and current_answer and line:
            current_explanation += " " + line
    if current_answer:
        questions.append((
            f"{current_question_num}.",
            current_answer,
            current_explanation or ""
        ))
        print(f"解析題目 {current_question_num}.: 答案={current_answer}, 解析={'有' if current_explanation else '無'}")
    
    return questions

def parse_with_fallback(text: str) -> List[Tuple[str, str, str]]:
    questions = []
    lines = text.split('\n')
    
    current_question = None
    current_answer = None
    current_explanation = None
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        if re.match(r'^\d+\.', line):
            if current_question:
                questions.append((
                    current_question, 
                    current_answer or "", 
                    current_explanation or ""
                ))
            
            current_question = line
            current_answer = None
            current_explanation = None
        
        elif re.search(r'答案\s*:\s*\([A-D]\)', line) and current_question:
            current_answer = line
        elif re.search(r'解析\s*:', line) and current_question:
            current_explanation = line
        elif current_explanation and current_question:
            current_explanation += " " + line
    if current_question:
        questions.append((
            current_question, 
            current_answer or "", 
            current_explanation or ""
        ))
    
    return questions

def main():
    # 測試檔案路徑
    doc_path = "7-2-解析卷簡.docx"
    
    if not os.path.exists(doc_path):
        print(f"錯誤: 找不到檔案 {doc_path}")
        return
    
    print("測試解析功能 - 從實際解析卷檔案讀取")
    print("=" * 60)
    
    questions = parse_source_document_test(doc_path)
    
    if not questions:
        print("未能解析到任何題目")
        return
    
    print(f"\n解析結果 (共 {len(questions)} 題):")
    print("=" * 60)
    
    # 顯示前 10 題或所有題目
    display_count = min(10, len(questions))
    for i, (question, answer, explanation) in enumerate(questions[:display_count], 1):
        print(f"第 {i} 題:")
        print(f"  題序: {question}")
        print(f"  答案: {answer}")
        if explanation and explanation.strip():
            print(f"  解析: {explanation}")
        else:
            print(f"  解析: (無解析)")
        print("-" * 40)
    
    if len(questions) > 10:
        print(f"... 還有 {len(questions) - 10} 題未顯示")

if __name__ == "__main__":
    main()

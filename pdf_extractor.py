"""
PDF 텍스트 추출기 - 두 가지 방법 비교
1. 직접 추출 (pdfplumber/PyMuPDF): PDF에 내장된 텍스트 레이어 추출
2. OCR 기반 추출 (pytesseract): 이미지로 변환 후 OCR 처리

사용법: python pdf_extractor.py
"""

import os
import sys
from pathlib import Path
from datetime import datetime

# PDF 처리 라이브러리
import fitz  # PyMuPDF
import pdfplumber

# 결과 저장 디렉토리
OUTPUT_DIR = Path("추출결과")
PDF_DIR = Path("참고자료")


def extract_with_pymupdf(pdf_path: Path) -> dict:
    """
    PyMuPDF를 사용하여 PDF에서 텍스트 직접 추출
    - 장점: 빠르고 정확 (텍스트 레이어가 있는 경우)
    - 단점: 이미지 PDF는 추출 불가
    """
    result = {
        "method": "PyMuPDF (직접 추출)",
        "text": "",
        "pages": [],
        "success": False,
        "error": None
    }
    
    try:
        doc = fitz.open(pdf_path)
        all_text = []
        
        for page_num, page in enumerate(doc):
            page_text = page.get_text("text")
            all_text.append(page_text)
            result["pages"].append({
                "page": page_num + 1,
                "text": page_text,
                "char_count": len(page_text)
            })
        
        result["text"] = "\n".join(all_text)
        result["success"] = True
        doc.close()
        
    except Exception as e:
        result["error"] = str(e)
    
    return result


def extract_with_pdfplumber(pdf_path: Path) -> dict:
    """
    pdfplumber를 사용하여 PDF에서 텍스트 직접 추출
    - 장점: 테이블 추출에 특화, 레이아웃 유지
    - 단점: 이미지 PDF는 추출 불가
    """
    result = {
        "method": "pdfplumber (직접 추출)",
        "text": "",
        "pages": [],
        "tables": [],
        "success": False,
        "error": None
    }
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = []
            
            for page_num, page in enumerate(pdf.pages):
                page_text = page.extract_text() or ""
                all_text.append(page_text)
                
                # 테이블 추출 시도
                tables = page.extract_tables()
                
                result["pages"].append({
                    "page": page_num + 1,
                    "text": page_text,
                    "char_count": len(page_text),
                    "table_count": len(tables)
                })
                
                if tables:
                    result["tables"].append({
                        "page": page_num + 1,
                        "tables": tables
                    })
            
            result["text"] = "\n".join(all_text)
            result["success"] = True
            
    except Exception as e:
        result["error"] = str(e)
    
    return result


def extract_with_ocr(pdf_path: Path) -> dict:
    """
    pytesseract를 사용하여 OCR 기반 텍스트 추출
    - 장점: 이미지 PDF도 처리 가능
    - 단점: 느리고, Tesseract 설치 필요
    
    참고: Tesseract가 설치되어 있어야 합니다.
    Windows: https://github.com/UB-Mannheim/tesseract/wiki
    """
    result = {
        "method": "OCR (pytesseract)",
        "text": "",
        "pages": [],
        "success": False,
        "error": None,
        "note": "Tesseract OCR이 설치되어 있어야 합니다."
    }
    
    try:
        import pytesseract
        from pdf2image import convert_from_path
        
        # Tesseract 경로 설정 (Windows)
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
        ]
        
        for path in tesseract_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break
        
        # PDF를 이미지로 변환
        # poppler가 필요합니다
        try:
            images = convert_from_path(pdf_path)
        except Exception as e:
            result["error"] = f"PDF를 이미지로 변환 실패 (poppler 필요): {str(e)}"
            return result
        
        all_text = []
        
        for page_num, image in enumerate(images):
            # OCR 수행 (한국어 + 영어 + 러시아어)
            page_text = pytesseract.image_to_string(
                image, 
                lang='kor+eng+rus',
                config='--psm 6'  # 균일한 텍스트 블록으로 간주
            )
            all_text.append(page_text)
            result["pages"].append({
                "page": page_num + 1,
                "text": page_text,
                "char_count": len(page_text)
            })
        
        result["text"] = "\n".join(all_text)
        result["success"] = True
        
    except ImportError as e:
        result["error"] = f"필요한 라이브러리가 없습니다: {str(e)}"
    except Exception as e:
        result["error"] = str(e)
    
    return result


def compare_methods(pdf_path: Path) -> dict:
    """
    세 가지 방법으로 추출한 결과를 비교
    """
    print(f"\n{'='*60}")
    print(f"파일: {pdf_path.name}")
    print(f"{'='*60}")
    
    results = {}
    
    # 방법 1: PyMuPDF
    print("  [1/3] PyMuPDF로 추출 중...")
    results["pymupdf"] = extract_with_pymupdf(pdf_path)
    
    # 방법 2: pdfplumber
    print("  [2/3] pdfplumber로 추출 중...")
    results["pdfplumber"] = extract_with_pdfplumber(pdf_path)
    
    # 방법 3: OCR (선택적)
    print("  [3/3] OCR 추출은 별도 확인 필요 (Tesseract/poppler 필요)")
    # OCR은 시간이 오래 걸리고 추가 설치가 필요하므로 기본적으로 비활성화
    # results["ocr"] = extract_with_ocr(pdf_path)
    
    # 결과 요약
    print("\n  결과 요약:")
    for method_name, result in results.items():
        char_count = len(result["text"])
        status = "✓ 성공" if result["success"] else f"✗ 실패: {result.get('error', 'Unknown')}"
        print(f"    - {result['method']}: {char_count:,}자 추출 {status}")
    
    return results


def save_results(pdf_name: str, results: dict, output_dir: Path):
    """
    추출 결과를 파일로 저장
    """
    output_dir.mkdir(exist_ok=True)
    
    # 파일명에서 확장자 제거
    base_name = pdf_name.replace(".pdf", "")
    
    for method_name, result in results.items():
        if result["success"] and result["text"].strip():
            output_file = output_dir / f"{base_name}_{method_name}.txt"
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(f"# 추출 방법: {result['method']}\n")
                f.write(f"# 원본 파일: {pdf_name}\n")
                f.write(f"# 추출 시간: {datetime.now().isoformat()}\n")
                f.write(f"# 총 문자 수: {len(result['text']):,}자\n")
                f.write(f"# 총 페이지: {len(result['pages'])}페이지\n")
                f.write("=" * 60 + "\n\n")
                f.write(result["text"])
            print(f"    저장됨: {output_file.name}")


def main():
    """
    메인 실행 함수
    """
    print("\n" + "=" * 60)
    print("PDF 텍스트 추출기 - 방법 비교")
    print("=" * 60)
    
    # PDF 파일 목록
    if not PDF_DIR.exists():
        print(f"오류: '{PDF_DIR}' 폴더가 없습니다.")
        return
    
    pdf_files = list(PDF_DIR.glob("*.pdf"))
    
    if not pdf_files:
        print(f"오류: '{PDF_DIR}' 폴더에 PDF 파일이 없습니다.")
        return
    
    print(f"\n총 {len(pdf_files)}개의 PDF 파일 발견")
    
    # 결과 저장 디렉토리 생성
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # 전체 결과 저장
    all_results = {}
    
    for pdf_file in pdf_files:
        results = compare_methods(pdf_file)
        all_results[pdf_file.name] = results
        save_results(pdf_file.name, results, OUTPUT_DIR)
    
    # 최종 요약
    print("\n" + "=" * 60)
    print("최종 요약")
    print("=" * 60)
    
    # 비교표 작성
    summary_file = OUTPUT_DIR / "비교_요약.txt"
    with open(summary_file, "w", encoding="utf-8") as f:
        f.write("PDF 텍스트 추출 결과 비교\n")
        f.write(f"추출 시간: {datetime.now().isoformat()}\n")
        f.write("=" * 80 + "\n\n")
        
        f.write(f"{'파일명':<50} {'PyMuPDF':>12} {'pdfplumber':>12}\n")
        f.write("-" * 80 + "\n")
        
        for pdf_name, results in all_results.items():
            pymupdf_chars = len(results.get("pymupdf", {}).get("text", ""))
            pdfplumber_chars = len(results.get("pdfplumber", {}).get("text", ""))
            
            # 파일명 줄이기
            short_name = pdf_name[:47] + "..." if len(pdf_name) > 50 else pdf_name
            f.write(f"{short_name:<50} {pymupdf_chars:>12,} {pdfplumber_chars:>12,}\n")
        
        f.write("\n\n주의사항:\n")
        f.write("- 문자 수가 0이면 해당 PDF에 텍스트 레이어가 없거나 이미지 PDF입니다.\n")
        f.write("- 이미지 PDF의 경우 OCR 방법을 사용해야 합니다.\n")
        f.write("- OCR 사용을 위해서는 Tesseract와 poppler 설치가 필요합니다.\n")
    
    print(f"\n결과가 '{OUTPUT_DIR}' 폴더에 저장되었습니다.")
    print(f"요약 파일: {summary_file}")


if __name__ == "__main__":
    main()

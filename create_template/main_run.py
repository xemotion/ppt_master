import os
import json
import sys
import logging
from datetime import datetime
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from typing import Set

# 현재 스크립트의 디렉토리를 Python 경로에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# 기존 모듈 import
from ppt_common import load_presentation, load_meta_info
from generate_meta import extract_meta_info
from create_template import update_slide


# 로깅 설정
logging.basicConfig(
    filename="main_process.log", 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# role 관리를 위한 전역 set
used_roles: Set[str] = set()

def generate_unique_role(base_role: str) -> str:
    """중복되지 않는 unique role을 생성합니다."""
    if base_role not in used_roles:
        used_roles.add(base_role)
        return base_role
    
    counter = 1
    while f"{base_role}_{counter}" in used_roles:
        counter += 1
    
    unique_role = f"{base_role}_{counter}"
    used_roles.add(unique_role)
    return unique_role

def process_pptx(input_pptx: str, output_dir: str):
    """
    입력된 PPTX 파일을 처리하여 메타데이터를 생성하고 텍스트를 업데이트합니다.
    
    이 과정은 다음과 같이 진행됩니다:
    1. 각 슬라이드별로 메타데이터를 추출하여 저장
    2. 각 슬라이드의 텍스트 업데이트
    3. 원본 파일 덮어쓰기
    """
    # 입출력 경로 확인
    input_path = Path(input_pptx)
    if not input_path.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {input_pptx}")
    
    # 출력 디렉토리 생성
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # 원본 프레젠테이션 로드
    prs = load_presentation(str(input_path))
    if not prs or not prs.slides:
        raise ValueError("유효하지 않거나 비어있는 프레젠테이션입니다.")
    
    # 파일 이름에서 확장자 제외한 부분 추출
    base_filename = input_path.stem
    meta_dir = output_path / f"{base_filename}_meta"
    meta_dir.mkdir(exist_ok=True)
    
    # 각 슬라이드 처리
    for slide_idx, slide in enumerate(prs.slides, 1):
        print(f"\n슬라이드 {slide_idx} 처리 중...")
        
        # 1. 메타데이터 추출 및 저장
        meta_info = extract_meta_info(slide, slide_idx, prs.slide_width, prs.slide_height)
        slide_meta_path = meta_dir / f"slide_{slide_idx}_meta.json"
        with open(slide_meta_path, 'w', encoding='utf-8') as f:
            json.dump(meta_info, f, ensure_ascii=False, indent=2)
        print(f"메타데이터 저장 완료: {slide_meta_path}")

        # meta_info = load_meta_info(slide_meta_path)
        # 2. 슬라이드 텍스트 업데이트
        update_slide(slide, meta_info, prs.slide_width, prs.slide_height)
        print(f"슬라이드 {slide_idx} 업데이트 완료")
    
    # 파일 저장하기
    update_output_path =f"{output_path}/result_{base_filename}.pptx" 
    prs.save(str(update_output_path))
    print(f"\nPPTX 업데이트 완료: {update_output_path}")
    
    return str(update_output_path)

def main():
    """
    메인 실행 함수
    입력 PPTX 파일을 처리하여 메타데이터를 생성하고 텍스트를 업데이트합니다.
    """
    file_list = os.listdir("your input directory")
    print(file_list)
    
    success = True
    for file_name in file_list:
        input_pptx = f"{input_dir}/{file_name}" 
        print(f"Start ------> {input_pptx}")
        output_dir = "output"
        
        try:
            output_path = process_pptx(input_pptx, output_dir)
            print(f"""
            처리가 완료되었습니다!
            출력 디렉토리: {output_path}
            """)
        except Exception as e:
            logger.error(f"오류 발생: {str(e)}", exc_info=True)
            print(f"오류 발생: {str(e)}")
            success = False
            continue
    
    return 0 if success else 1


if __name__ == "__main__":
    main() 

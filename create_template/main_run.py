import os
import json
import logging
import argparse
from datetime import datetime
from pathlib import Path
from pptx import Presentation

# 기존 모듈 import
from generate_meta_info import extract_meta_info
from create_template import load_presentation, update_slide

# 로깅 설정
logging.basicConfig(
    filename="main_process.log", 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def process_pptx(input_pptx: str, output_dir: str):
    """
    입력된 PPTX 파일을 처리하여 메타데이터를 생성하고 각 슬라이드별로 처리하여 저장합니다.
    
    이 과정은 다음과 같이 진행됩니다:
    1. 입력 PPTX 파일에서 메타데이터를 추출합니다 (generate_meta_info.py)
       - 일반 텍스트 요소, 표 셀, 그룹화된 요소(중첩 그룹 포함) 모두 처리
    2. 각 슬라이드를 개별적으로 처리합니다 (create_template.py)
       - 슬라이드 별로 개별 PPTX 파일 생성
       - 해당 슬라이드만 처리하여 저장
    
    Args:
        input_pptx: 입력 PPTX 파일 경로
        output_dir: 출력 디렉토리
        
    Returns:
        str: 출력 디렉토리 경로
    """
    # 입출력 경로 확인
    input_path = Path(input_pptx)
    if not input_path.exists():
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {input_pptx}")
    
    # 출력 디렉토리 생성
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # 메타데이터 파일 경로
    meta_filename = f"{input_path.stem}_meta_info.json"
    meta_path = output_path / meta_filename
    
    logger.info(f"1. 메타데이터 생성 시작: {input_path} -> {meta_path}")
    print(f"1. 메타데이터 생성 중: {input_path.name}...")
    
    # 메타데이터 추출
    extract_meta_info(str(input_path), str(meta_path))
    
    logger.info(f"2. 메타데이터 생성 완료: {meta_path}")
    print(f"2. 메타데이터 생성 완료: {meta_path.name}")
    
    # 메타데이터 로드
    with open(meta_path, 'r', encoding='utf-8') as f:
        template_schema = json.load(f)
    
    # 원본 프레젠테이션 로드
    prs = load_presentation(str(input_path))
    if not prs or not prs.slides:
        raise ValueError("유효하지 않거나 비어있는 프레젠테이션 템플릿입니다.")
    
    # 각 슬라이드를 개별적으로 처리
    for slide_idx, slide in enumerate(prs.slides):
        # 새 프레젠테이션 객체 생성 (슬라이드당 하나)
        single_slide_prs = Presentation(str(input_path))
        
        # 처리할 슬라이드 이외의 슬라이드 삭제
        slides_to_delete = list(range(len(single_slide_prs.slides)))
        slides_to_delete.remove(slide_idx)  # 유지할 슬라이드 인덱스 제외
        
        # 슬라이드 삭제 (뒤에서부터 삭제해야 인덱스가 변경되지 않음)
        for idx in sorted(slides_to_delete, reverse=True):
            rId = single_slide_prs.slides._sldIdLst[idx].rId
            single_slide_prs.part.drop_rel(rId)
            del single_slide_prs.slides._sldIdLst[idx]
        
        logger.info(f"3. 슬라이드 {slide_idx+1} 처리 중...")
        print(f"3. 슬라이드 {slide_idx+1} 처리 중...")
        
        # 단일 슬라이드 업데이트
        update_slide(single_slide_prs.slides[0], template_schema)
        
        # 결과 저장
        output_filename = f"{input_path.stem}_{slide_idx+1}.pptx"
        output_file_path = output_path / output_filename
        single_slide_prs.save(str(output_file_path))
        
        logger.info(f"4. 슬라이드 {slide_idx+1} 저장 완료: {output_file_path}")
        print(f"4. 슬라이드 {slide_idx+1} 저장 완료: {output_filename}")
    
    logger.info(f"모든 슬라이드 처리 완료. 결과 디렉토리: {output_path}")
    print(f"모든 슬라이드 처리 완료. 결과 디렉토리: {output_path}")
    return str(output_path)

def main():
    """
    메인 함수 - 명령줄 인자를 처리하고 PPTX 처리 실행
    
    사용 예:
    python main_process.py --input "input/presentation.pptx" --output "output"
    또는
    python main_process.py -i "input/presentation.pptx" -o "output"
    """
    parser = argparse.ArgumentParser(description='PowerPoint 템플릿 처리 도구')
    parser.add_argument('--input', '-i', required=True, help='입력 PPTX 파일 경로')
    parser.add_argument('--output', '-o', default='./output', help='출력 디렉토리 경로 (기본값: ./output)')
    
    args = parser.parse_args()
    
    try:
        output_path = process_pptx(args.input, args.output)
        print(f"""
            처리가 완료되었습니다!

            - 메타데이터 JSON 파일: {os.path.join(args.output, Path(args.input).stem + '_meta_info.json')}
            - 처리된 슬라이드: {output_path} 디렉토리 내 파일들

            참고: 
            * 모든 슬라이드는 개별 파일로 분리되었습니다.
            * 그룹화된 요소 및 표 셀도 처리되었습니다.
            """)
    except Exception as e:
        logger.error(f"오류 발생: {str(e)}", exc_info=True)
        print(f"오류 발생: {str(e)}")
        return 1
        
    return 0

if __name__ == "__main__":
    main() 

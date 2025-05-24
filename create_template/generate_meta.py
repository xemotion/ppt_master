import json
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from dotenv import load_dotenv
from openai import AzureOpenAI
from typing import Dict, Any, List, Optional
from pathlib import Path
import logging

# 공통 모듈에서 함수 가져오기
from ppt_common import (
    get_shape_position,
    get_type_info,
    make_element_id,
    is_tag_or_label,
    logger,
    find_tag_element,
    generate_position_key,
    is_tag_identifier
)

# 로깅 설정
logging.basicConfig(
    filename="ppt_process.log",
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def determine_element_role(shape, pos, slide_number, slide_width, slide_height, shape_type):
    """
    슬라이드 내 요소의 위치와 유형을 고려하여 의미있는 역할 이름을 생성합니다.
    
    Args:
        shape: PPT 도형 객체
        pos: 위치 정보 딕셔너리 (left_percent, top_percent, width_percent, height_percent)
        slide_number: 슬라이드 번호 (함수 내부에서는 사용하지 않음)
        slide_width: 슬라이드 너비
        slide_height: 슬라이드 높이
        shape_type: 도형 유형 (TEXT_BOX, AUTO_SHAPE 등)
        
    Returns:
        str: 요소의 역할 이름
    """
    # 위치 결정 (상단, 중간, 하단)
    if pos["top_percent"] < 20:
        vertical_pos = "top"
    elif pos["top_percent"] > 70:
        vertical_pos = "bottom"
    else:
        vertical_pos = "middle"
    
    # 위치 결정 (좌측, 중앙, 우측)
    if pos["left_percent"] < 30:
        horizontal_pos = "left"
    elif pos["left_percent"] > 60:
        horizontal_pos = "right"
    else:
        horizontal_pos = "center"
    
    # 크기 결정 (크기에 따라 중요도 부여)
    if pos["width_percent"] > 70 or pos["height_percent"] > 30:
        size = "main"
    elif pos["width_percent"] > 40 or pos["height_percent"] > 15:
        size = "sub"
    else:
        size = "detail"
    
    # 요소 유형에 따른 기능적 역할
    if shape_type == "TEXT_BOX":
        # 텍스트 크기와 위치로 역할 추론
        if vertical_pos == "top" and (size == "main" or size == "sub"):
            functional_role = "title"
        elif vertical_pos == "top" and size == "detail":
            functional_role = "header"
        elif vertical_pos == "bottom":
            functional_role = "footer"
        else:
            functional_role = "content"
    elif shape_type == "AUTO_SHAPE":
        functional_role = "shape"
    elif shape_type == "PICTURE":
        functional_role = "image"
    elif shape_type == "GROUP":
        functional_role = "group"
    elif shape_type == "TABLE":
        functional_role = "table"
    elif shape_type == "CHART":
        functional_role = "chart"
    else:
        functional_role = "element"
    
    # 특수 케이스 처리: 상단 제목
    if vertical_pos == "top" and horizontal_pos == "center" and size == "main" and shape_type == "TEXT_BOX":
        return f"{vertical_pos}_{horizontal_pos}_main_title"
    
    # 일반적인 이름 생성 (슬라이드 번호 제외)
    return f"{vertical_pos}_{horizontal_pos}_{size}_{functional_role}"

def call_llm_for_meta(text, pos, type_name, slide_number, slide_width, slide_height, client, deployment_name):
    """
    텍스트 요소에 대한 역할과 설명을 생성합니다.
    LLM API를 호출하거나 규칙 기반으로 역할을 생성합니다.
    
    Args:
        text: 텍스트 내용
        pos: 위치 정보
        type_name: 도형 유형
        slide_number: 슬라이드 번호
        slide_width: 슬라이드 너비
        slide_height: 슬라이드 높이
        client: OpenAI 클라이언트
        deployment_name: 배포 이름
        
    Returns:
        tuple: (역할 이름, 설명) 튜플
    """
    try:
        # 태그/라벨 확인
        is_tag, tag_type = is_tag_or_label(text, pos, type_name)
        
        # 태그/라벨로 판단되면 특별 처리
        if is_tag:
            print(f"태그/라벨 요소 감지: '{text}' (유형: {tag_type})")
            
            # 태그/라벨 역할 이름 생성
            tag_count = getattr(call_llm_for_meta, f"{tag_type}_count", 0) + 1
            setattr(call_llm_for_meta, f"{tag_type}_count", tag_count)
            
            role = f"{tag_type}_{tag_count}"
            description = f"슬라이드 {slide_number}의 {tag_type} 요소 (텍스트: '{text}')"
            
            return role, description
        
        # 원본 텍스트 길이 계산 (나중에 제한으로 사용)
        original_text_length = len(text)
        max_role_length = min(original_text_length, 50)  # 역할 이름은 최대 50자 또는 원본 텍스트 길이 중 작은 값
       
        # 텍스트가 숫자로만 되어 있는지 확인
        is_number_only = text.isdigit()
        
        # 텍스트가 1-2글자 특수기호인지 확인
        is_special_char = len(text) <= 2 and not text.isalnum()
        
        # 숫자만 있거나 1-2글자 특수기호인 경우 원본 텍스트를 role로 사용
        if is_number_only or is_special_char:
            return text, f"슬라이드 {slide_number}의 특수 요소 (위치: {pos['left_percent']:.1f}%, {pos['top_percent']:.1f}%)"
        
        # 의미있는 역할 이름 생성
        role_name = determine_element_role(None, pos, slide_number, slide_width, slide_height, type_name)
        
        # 역할 이름 길이 제한
        if len(role_name) > max_role_length:
            role_name = role_name[:max_role_length]
            print(f"역할 이름 길이 제한: '{role_name}'")
        
        # LLM API 호출 여부 확인
        use_llm = True  # LLM 호출 활성화
        
        if use_llm and all([client, deployment_name]):
            system_prompt = f"""
                You are an expert in slide design and template structure analysis.

                You will receive a text element from a presentation slide along with its position and shape information.

                Your task is to:
                1. Assign a meaningful, **snake_case** role name (`role`) based on its visual function and position. Example: `main_title`, `left_detail_body_text`, `top_right_summary`, `comparison_title_right_1`
                - If there are multiple similar elements, append `_1`, `_2`, etc. to make each role unique.
                - DO NOT use unknown, misc, or generic labels.
                - VERY IMPORTANT: Keep the role name within {max_role_length} characters.
                
                2. Write a structural `description` that clearly explains the **role and functional purpose of this element in the overall slide template**.
                - DO NOT mention the actual content of the text.
                - DO NOT describe the topic (e.g. SW, DB, AI).
                - Describe only the visual structure, layout function, importance level, and placement logic.
                - Example: "This element is placed at the center top of the slide and serves as the main heading, establishing the visual hierarchy of the architecture overview template."

                Respond with a JSON object:
                {{
                "role": "...",
                "description": "..."
                }}
                Respond ONLY with a valid JSON. Do not include any markdown, bullet points, or extra commentary.
                """.strip()

            user_prompt = {
                "text": text,
                "position": pos,
                "type": type_name,
                "slide_number": slide_number,
                "max_role_length": max_role_length,
            }
            
            try:
                print(f"LLM API 호출 중... (슬라이드 {slide_number}, 텍스트: '{text[:30]}{'...' if len(text) > 30 else ''}')")
                response = client.chat.completions.create(
                    model=deployment_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": json.dumps(user_prompt, ensure_ascii=False, indent=2)}
                    ],
                    temperature=0.7,
                    max_tokens=300
                )
                response_text = response.choices[0].message.content.strip()
                print(f"LLM Raw Response (슬라이드 {slide_number}):", response_text)  # 응답 확인

                # JSON 형식이 마크다운 코드 블록으로 감싸져 있는 경우 처리
                if response_text.startswith("```"):
                    response_text = response_text.split("```")[1].strip()
                    if response_text.startswith("json"):
                        response_text = response_text[4:].strip()

                try:
                    data = json.loads(response_text)
                    role = data["role"]
                    description = data["description"]
                    
                    # 길이 제한 강제 적용
                    if len(role) > max_role_length:
                        role = role[:max_role_length]
                        print(f"LLM 응답 역할 이름 길이 초과, 잘림: '{role}'")
                    

                    return role, description
                except Exception as e:
                    print(f"[Parse Error] {e}")
                    print(f"[RAW Response] {response_text}")
                    return role_name, f"JSON 파싱 오류: {str(e)[:100]}"
            except Exception as e:
                print(f"[LLM API Error] {e}")
                return role_name, f"LLM API 오류: {str(e)[:100]}"
        
        # LLM을 사용하지 않거나 호출에 실패한 경우 의미 있는 description 생성
        vertical_pos = "상단" if pos["top_percent"] < 20 else "하단" if pos["top_percent"] > 70 else "중간"
        horizontal_pos = "좌측" if pos["left_percent"] < 30 else "우측" if pos["left_percent"] > 60 else "중앙"
        size_desc = "주요" if pos["width_percent"] > 70 or pos["height_percent"] > 30 else "보조" if pos["width_percent"] > 40 or pos["height_percent"] > 15 else "세부"
        
        type_desc = {
            "TEXT_BOX": "텍스트",
            "AUTO_SHAPE": "도형",
            "PICTURE": "이미지",
            "GROUP": "그룹",
            "TABLE": "표",
            "CHART": "차트"
        }.get(type_name, "요소")
        
        description = f"슬라이드 {slide_number}의 {vertical_pos} {horizontal_pos}에 위치한 {size_desc} {type_desc} 요소"
        

        
        return role_name, description
    except Exception as e:
        # 예상치 못한 오류 발생 시 안전하게 처리
        print(f"[Unexpected Error in call_llm_for_meta] {e}")
        return "unknown", f"오류: {str(e)[:100]}"
        
# 태그/라벨 카운터 초기화
call_llm_for_meta.tag_count = 0
call_llm_for_meta.label_count = 0
call_llm_for_meta.cic_label_count = 0
call_llm_for_meta.ui_element_count = 0

def generate_unique_key(text: str, position: dict) -> str:
    """텍스트와 위치 정보를 조합하여 고유 키 생성"""
    # 위치를 5% 단위로 반올림하여 근접 위치는 같은 키로 처리
    left_group = round(position['left_percent'] / 5) * 5
    top_group = round(position['top_percent'] / 5) * 5
    return f"{text}__pos_{left_group}_{top_group}"

def extract_text_from_shape(shape) -> str:
    """도형에서 텍스트를 추출합니다."""
    if not hasattr(shape, "text_frame"):
        return ""
        
    paragraphs = [p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)

def extract_table_info(shape) -> Optional[Dict[str, Any]]:
    """표에서 정보를 추출합니다."""
    if shape.shape_type != MSO_SHAPE_TYPE.TABLE:
        return None
        
    table_data = []
    for row_idx, row in enumerate(shape.table.rows, 1):
        row_data = []
        for col_idx, cell in enumerate(row.cells, 1):
            cell_text = "\n".join([p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()])
            if cell_text:
                row_data.append({
                    "text": cell_text,
                    "row": row_idx,
                    "col": col_idx
                })
        if row_data:
            table_data.extend(row_data)
            
    return table_data if table_data else None

def process_shape(shape, slide_number: int, slide_width: int, slide_height: int) -> Optional[Dict[str, Any]]:
    """개별 도형을 처리하고 메타 정보를 추출합니다."""
    # 기본 위치 정보 추출
    pos = get_shape_position(shape, slide_width, slide_height)
    type_name = get_type_info(shape)
    
    # 표 처리
    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        table_info = extract_table_info(shape)
        if table_info:
            return {
                "type": "table",
                "position": pos,
                "cells": table_info
            }
        return None
    
    # 텍스트 추출
    text = extract_text_from_shape(shape)
    if not text:
        return None
        
    # 위치 기반 키 생성
    position_key = generate_position_key(text, pos)
    
    # 메타 정보 구성
    return {
        "id": make_element_id(slide_number, pos, type_name),
        "type": type_name,
        "text": text,
        "position": pos,
        "position_key": position_key
    }

def extract_meta_info(slide, slide_number: int, slide_width: int, slide_height: int) -> Dict[str, Any]:
    """슬라이드에서 메타 정보를 추출합니다."""
    meta_info = {
        "slide_number": slide_number,
        "fields": {}
    }
    
    # 모든 도형 처리
    text_counts = {}  # 텍스트 중복 카운트용
    
    for shape in slide.shapes:
        shape_info = process_shape(shape, slide_number, slide_width, slide_height)
        if not shape_info:
            continue
            
        if shape_info["type"] == "table":
            # 표 셀 처리
            for cell in shape_info["cells"]:
                cell_text = cell["text"]
                field_key = cell_text
                
                # 중복 텍스트 처리
                if cell_text in text_counts:
                    text_counts[cell_text] += 1
                    field_key = f"{cell_text}_{text_counts[cell_text]}"
                else:
                    text_counts[cell_text] = 1
                
                meta_info["fields"][field_key] = {
                    "type": "table_cell",
                    "original_text": cell_text,
                    "table_info": {
                        "row": cell["row"],
                        "col": cell["col"]
                    },
                    "text_count": text_counts[cell_text]
                }
        else:
            # 일반 도형 처리
            text = shape_info["text"]
            position_key = shape_info["position_key"]
            
            # 위치 기반 키로 저장
            meta_info["fields"][position_key] = {
                "type": shape_info["type"],
                "original_text": text,
                "position": shape_info["position"],
                "element_id": shape_info["id"]
            }
            
            # 태그/라벨 요소 처리
            if is_tag_identifier(text):
                meta_info["fields"][position_key]["is_tag"] = True
    
    return meta_info

def process_presentation(pptx_path: str) -> None:
    """프레젠테이션 파일을 처리하고 메타 정보를 저장합니다."""
    # 입력 파일 확인
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")
        
    # 프레젠테이션 로드
    prs = Presentation(pptx_path)
    if not prs or not prs.slides:
        raise ValueError("Invalid or empty presentation")
        
    # 출력 디렉토리 설정
    input_stem = Path(pptx_path).stem
    output_dir = f"output/{input_stem}"
    os.makedirs(output_dir, exist_ok=True)
    
    # 각 슬라이드 처리
    for slide_number, slide in enumerate(prs.slides, 1):
        # 메타 정보 추출
        meta_info = extract_meta_info(slide, slide_number, prs.slide_width, prs.slide_height)
        
        # 메타 정보 저장
        output_path = os.path.join(output_dir, f"{input_stem}_slide_{slide_number}_meta.json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(meta_info, f, ensure_ascii=False, indent=2)
            
        logger.info(f"슬라이드 {slide_number} 메타 정보 저장 완료: {output_path}")
        print(f"슬라이드 {slide_number} 메타 정보 저장 완료: {output_path}")
    
    print(f"모든 슬라이드의 메타 정보 저장 완료. 폴더: {output_dir}")

if __name__ == "__main__":
    # 테스트용 PPTX 파일 경로
    test_pptx = "input/test.pptx"
    process_presentation(test_pptx) 

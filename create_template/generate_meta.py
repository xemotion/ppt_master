import json
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from dotenv import load_dotenv
from openai import AzureOpenAI
from typing import Dict, Any

# 공통 모듈에서 함수 가져오기
from ppt_common import (
    get_shape_position,
    get_type_info,
    make_element_id,
    is_tag_or_label,
    logger
)

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
        max_desc_length = original_text_length * 2  # 설명은 원본 텍스트 길이의 2배로 제한
        
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
                - VERY IMPORTANT: Keep the role name within max 15 characters.
                
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
                "max_desc_length": max_desc_length
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
                    
                    if len(description) > max_desc_length:
                        description = description[:max_desc_length]
                        print(f"LLM 응답 설명 길이 초과, 잘림: '{description[:50]}...'")
                    
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
        
        # 설명 길이 제한
        if len(description) > max_desc_length:
            description = description[:max_desc_length]
        
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

def process_shape(shape, slide_number, slide_width, slide_height, meta, text_counter, role_counters, client, deployment_name, group_context=None):
    """개별 도형을 처리하는 함수 (추출 로직을 분리하여 재사용성 높임)"""
    # 그룹 도형 처리 (재귀적으로 중첩된 그룹까지 처리)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        group_id = group_context or "그룹"
        print(f"그룹 요소 처리 중: {group_id}")
        try:
            for idx, shape_in_group in enumerate(shape.shapes):
                # 그룹 내에 또 다른 그룹이 있으면 재귀적으로 처리
                if shape_in_group.shape_type == MSO_SHAPE_TYPE.GROUP:
                    nested_group_context = f"{group_id}_중첩그룹{idx+1}"
                    process_shape(shape_in_group, slide_number, slide_width, slide_height,
                                 meta, text_counter, role_counters, client, deployment_name,
                                 group_context=nested_group_context)
                else:
                    # 일반 도형 처리
                    process_shape(shape_in_group, slide_number, slide_width, slide_height,
                                 meta, text_counter, role_counters, client, deployment_name,
                                 group_context=f"{group_id}_{idx+1}")
            return
        except Exception as e:
            print(f"그룹 요소 처리 중 오류: {e}")
            # 오류가 발생해도 계속 진행 (다른 방식으로 처리 시도)
    
    # 표 요소 처리
    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        print(f"표 요소 발견 (행: {len(shape.table.rows)}, 열: {len(shape.table.columns)})")
        try:
            for row_idx, row in enumerate(shape.table.rows):
                for col_idx, cell in enumerate(row.cells):
                    cell_text = "\n".join([p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()])
                    if not cell_text:
                        continue
                        
                    print(f"표 셀 텍스트 추출: R{row_idx+1}C{col_idx+1} - '{cell_text[:30]}{'...' if len(cell_text) > 30 else ''}'")
                    
                    # 위치 정보 계산 (대략적 추정)
                    pos = {
                        "left_percent": round(((shape.left + (col_idx * shape.width / len(shape.table.columns))) / slide_width) * 100, 2),
                        "top_percent": round(((shape.top + (row_idx * shape.height / len(shape.table.rows))) / slide_height) * 100, 2),
                        "width_percent": round((shape.width / len(shape.table.columns) / slide_width) * 100, 2),
                        "height_percent": round((shape.height / len(shape.table.rows) / slide_height) * 100, 2)
                    }
                    
                    type_name = "TABLE_CELL"
                    element_id = make_element_id(slide_number, pos, type_name)
                    
                    # 역할 이름 생성
                    base_role, description = call_llm_for_meta(cell_text, pos, type_name, slide_number, slide_width, slide_height, client, deployment_name)
                    
                    # 표 셀 컨텍스트 포함
                    base_role = f"{base_role}_table_{row_idx+1}_{col_idx+1}"
                    
                    # 중복 텍스트 처리 - 텍스트 내용 자체를 키로 사용
                    text_key = cell_text
                    
                    # 텍스트 내용이 이미 처리된 적 있으면 카운터 증가
                    if text_key in text_counter:
                        text_counter[text_key] += 1
                        # 텍스트 키는 원본 텍스트 자체에 번호 붙이기
                        numbered_key = f"{text_key}_{text_counter[text_key]}"
                        # 역할 이름에도 번호 붙이기
                        role = f"{base_role}_{text_counter[text_key]}"
                    else:
                        text_counter[text_key] = 1
                        numbered_key = text_key
                        # 첫 번째 발견에도 1을 붙여 일관성 유지
                        role = f"{base_role}_1"
                    
                    # 메타데이터 저장
                    meta["fields"][numbered_key] = {
                        "role": role,
                        "role_description": description,
                        "element_id": element_id,
                        "slide_number": slide_number,
                        "position": pos,
                        "type": type_name,
                        "text_count": text_counter[text_key],  # 동일 텍스트의 몇 번째 등장인지 저장
                        "table_info": {
                            "row": row_idx + 1,
                            "col": col_idx + 1,
                            "total_rows": len(shape.table.rows),
                            "total_cols": len(shape.table.columns)
                        }
                    }
        except Exception as e:
            print(f"표 처리 중 오류: {e}")
        return
    
    # 텍스트 프레임이 없는 경우 처리 중단
    if not hasattr(shape, "text_frame") or not shape.text_frame:
        return
            
    # 전체 텍스트 프레임의 내용을 하나의 키로 추출
    # 단, 각 단락은 보존
    shape_text = "\n".join([p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()])
    
    if not shape_text:
        return
    
    # 컨텍스트 정보 추가 (그룹 내부 요소인 경우)
    context_info = f"_{group_context}" if group_context else ""
    
    # 위치 정보 및 유형 정보 가져오기
    pos = get_shape_position(shape, slide_width, slide_height)
    type_name = get_type_info(shape)
    element_id = make_element_id(slide_number, pos, type_name)
    
    # 역할 이름 생성
    base_role, description = call_llm_for_meta(shape_text, pos, type_name, slide_number, slide_width, slide_height, client, deployment_name)
    
    # 그룹 컨텍스트 포함
    if group_context:
        base_role = f"{base_role}_{group_context}"
    
    # 중복 텍스트 처리 - 텍스트 내용 자체를 키로 사용 (컨텍스트도 포함)
    text_key = f"{shape_text}{context_info}"
    
    # 텍스트 내용이 이미 처리된 적 있으면 카운터 증가
    if text_key in text_counter:
        text_counter[text_key] += 1
        # 텍스트 키는 원본 텍스트 자체에 번호 붙이기
        numbered_key = f"{text_key}_{text_counter[text_key]}"
        # 역할 이름에도 번호 붙이기
        role = f"{base_role}_{text_counter[text_key]}"
    else:
        text_counter[text_key] = 1
        numbered_key = text_key
        # 첫 번째 발견에도 1을 붙여 일관성 유지
        role = f"{base_role}_1"
    
    # 디버깅을 위해 추가 정보 출력
    print(f"추출된 텍스트: '{shape_text[:50]}{'...' if len(shape_text) > 50 else ''}'")
    print(f"생성된 역할명: '{role}' (카운트: {text_counter[text_key]})")
    print("---")
    
    # 메타데이터 저장
    meta["fields"][numbered_key] = {
        "role": role,
        "role_description": description,
        "element_id": element_id,
        "slide_number": slide_number,
        "position": pos,
        "type": type_name,
        "text_count": text_counter[text_key],  # 동일 텍스트의 몇 번째 등장인지 저장
        "in_group": True if group_context else False,
        "group_context": group_context if group_context else None
    }

def extract_meta_info(pptx_path: str, meta_path: str):
    load_dotenv()
    client = AzureOpenAI(
        api_key=os.getenv("AZURE_OPENAI_API_KEY"),
        api_version=os.getenv("AZURE_OPENAI_VERSION"),
        azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
    )
    deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")
    prs = Presentation(pptx_path)
    meta = {"fields": {}}
    text_counter = {}
    
    # 역할 이름 카운터 추가 (고유성 보장)
    role_counters = {}
    
    for slide_number, slide in enumerate(prs.slides, 1):
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        print(f"\n===== 슬라이드 {slide_number} 처리 중 =====")
        
        for shape in slide.shapes:
            # 모든 도형은 process_shape로 통합 처리 (그룹 포함)
            process_shape(shape, slide_number, slide_width, slide_height, 
                          meta, text_counter, role_counters, client, deployment_name)
            
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print(f"\n메타 정보가 {meta_path}에 저장되었습니다.")

def main():
    INPUT_DIR ="{input_diretory_path}"
    OUTPUT_DIR ="{output_directory_path}"
    FILE_NAME ="2"
    pptx_path = f"{INPUT_DIR}/{FILE_NAME}.pptx"
    meta_path = f"{OUTPUT_DIR}/meta_{FILE_NAME}.json"
    extract_meta_info(pptx_path, meta_path)

if __name__ == "__main__":
    main() 

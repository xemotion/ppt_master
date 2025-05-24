import os
import json
from datetime import datetime
from typing import Dict, List, Optional, Union, Any, Tuple
import subprocess
import logging
from pathlib import Path
from dataclasses import dataclass

import pandas as pd
from pptx import Presentation
from dotenv import load_dotenv
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE_TYPE

# 공통 모듈에서 함수 가져오기
from ppt_common import (
    change_text_to, 
    is_tag_identifier, 
    find_tag_element,
    find_shape_by_text_with_count,
    extract_count_from_field_name,
    logger
)

# Configure logging
logging.basicConfig(
    filename="logging_test.log", 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def load_presentation(template_path: str) -> Optional[Presentation]:
    """
    Load a PowerPoint presentation.
    Args:
        template_path: Path to the template file
    Returns:
        Presentation object, or None if loading fails
    """
    try:
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        return Presentation(template_path)

    except Exception as e:
        logger.error(f"Error loading presentation: {str(e)}")
        return None

def process_tag_element(slide, field_name, content, normalized_field_name):
    """
    태그/라벨 요소를 처리합니다.
    
    Args:
        slide: 처리할 슬라이드
        field_name: 필드 이름
        content: 업데이트할 내용
        normalized_field_name: 정규화된 필드 이름
        
    Returns:
        bool: 처리 성공 여부
    """
    # 태그/라벨 요소 찾기
    tag_shape = find_tag_element(slide, normalized_field_name)
    
    if tag_shape:
        logger.info(f"태그/라벨 요소 발견: '{field_name}' -> '{content}'")
        change_text_to(tag_shape, str(content))
        return True
        
    return False

def update_table_cell(slide, field_name: str, content: str, table_info: dict, normalized_field_name: str) -> bool:
    """표 셀의 텍스트를 업데이트합니다."""
    row = table_info.get('row', 0)
    col = table_info.get('col', 0)
    
    # 필드 이름에서 카운트 추출
    clean_field_name, clean_normalized_field_name, expected_count = extract_count_from_field_name(field_name)
    
    # 기존 text_count 정보가 있는 경우 활용
    meta_json = getattr(update_table_cell, 'meta_json', None)
    if meta_json and field_name in meta_json.get('fields', {}):
        from_meta_count = meta_json['fields'][field_name].get('text_count', None)
        if from_meta_count:
            expected_count = from_meta_count
    
    # 표 인스턴스 카운터
    current_count = 0
    
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            try:
                # 표 행렬이 충분히 큰지 확인
                if len(shape.table.rows) >= row and len(shape.table.columns) >= col:
                    # 1부터 시작하므로 인덱스는 1 빼기
                    cell = shape.table.cell(row-1, col-1)
                    cell_text = "\n".join([p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()])
                    
                    if not cell_text:
                        continue
                        
                    # 셀 텍스트 정규화
                    normalized_cell_text = ''.join(cell_text.lower().split())
                    
                    # 텍스트 일치 확인
                    if normalized_cell_text == clean_normalized_field_name or cell_text.lower() == clean_field_name.lower():
                        current_count += 1
                        
                        # 원하는 순번의 인스턴스를 찾았는지 확인
                        if current_count == expected_count:
                            logger.info(f"표 셀 일치 발견 [행:{row}, 열:{col}, 인스턴스 #{expected_count}]: '{cell_text[:30]}...' -> '{content[:30]}...'")
                            change_text_to(cell, str(content))
                            return True
            except Exception as e:
                logger.error(f"표 셀 업데이트 중 오류: {e}")
    
    if current_count > 0:
        logger.info(f"경고: 요청한 표 셀을 {current_count}개 찾았지만, 요청한 {expected_count}번째 인스턴스는 찾지 못했습니다.")
    
    return False

def process_regular_shapes(slide, field_name: str, content: str, normalized_field_name: str, expected_count=1) -> bool:
    """일반 도형들의 텍스트를 처리합니다."""
    # 필드 이름에서 카운트 추출
    clean_field_name, clean_normalized_field_name, count_from_name = extract_count_from_field_name(field_name)
    
    # 명시적으로 지정된 expected_count가 없으면 필드 이름에서 추출한 count 사용
    if expected_count == 1 and count_from_name > 1:
        expected_count = count_from_name
    
    # 기존 text_count 정보가 있는 경우 활용
    from_meta_count = None
    meta_json = getattr(process_regular_shapes, 'meta_json', None)
    if meta_json and field_name in meta_json.get('fields', {}):
        from_meta_count = meta_json['fields'][field_name].get('text_count', None)
        if from_meta_count:
            expected_count = from_meta_count
    
    # 번호 기반으로 해당 요소 찾기
    shape, found_count = find_shape_by_text_with_count(slide, clean_field_name, clean_normalized_field_name, expected_count)
    
    if shape:
        logger.info(f"텍스트 일치 발견 (인스턴스 #{expected_count}): '{clean_field_name[:30]}...' -> '{content[:30]}...'")
        change_text_to(shape, str(content))
        return True
            
    # 여러 인스턴스 중 하나를 못 찾았어도 다른 인스턴스가 있는지 확인
    if found_count > 0:
        logger.info(f"경고: '{clean_field_name}' 텍스트를 가진 요소를 {found_count}개 찾았지만, 요청한 {expected_count}번째 인스턴스는 찾지 못했습니다.")
        
    return False

def process_GROUPs(slide, field_name: str, content: str, normalized_field_name: str) -> bool:
    """그룹 요소 내부의 도형들을 처리합니다. 중첩 그룹도 처리합니다."""
    found_match = False
    
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                # 그룹 내부 처리를 재귀적으로 수행하는 내부 함수
                found_match = process_group_recursive(shape, field_name, content, normalized_field_name)
                if found_match:
                    return True
            except Exception as e:
                logger.error(f"그룹 요소 처리 중 오류: {e}")
    
    return False

def process_group_recursive(GROUP, field_name: str, content: str, normalized_field_name: str, depth: int = 1, expected_count: int = 1) -> bool:
    """그룹 내부를 재귀적으로 처리하는 함수"""
    # 필드 이름에서 카운트 추출
    clean_field_name, clean_normalized_field_name, count_from_name = extract_count_from_field_name(field_name)
    
    # 명시적으로 지정된 expected_count가 없으면 필드 이름에서 추출한 count 사용
    if expected_count == 1 and count_from_name > 1:
        expected_count = count_from_name
    
    # 현재 그룹 내 일치하는 요소 카운터
    current_count = 0
    
    # 그룹 내의 모든 도형을 처리
    for shape_in_group in GROUP.shapes:
        # 중첩 그룹 처리
        if shape_in_group.shape_type == MSO_SHAPE_TYPE.GROUP:
            logger.info(f"중첩 그룹 발견 (깊이: {depth})")
            found = process_group_recursive(shape_in_group, clean_field_name, content, clean_normalized_field_name, depth + 1, expected_count)
            if found:
                return True
        
        # 텍스트 프레임이 있는 도형 처리
        if hasattr(shape_in_group, "text_frame") and shape_in_group.text_frame:
            # 그룹 내 도형 텍스트 추출
            shape_text = "\n".join([p.text.strip() for p in shape_in_group.text_frame.paragraphs if p.text.strip()])
            
            if not shape_text:
                continue
            
            # 텍스트 정규화
            normalized_shape_text = ''.join(shape_text.lower().split())
            
            # 텍스트 비교 - 정규화된 텍스트 또는 원본 텍스트 일치 확인
            if normalized_shape_text == clean_normalized_field_name or shape_text.lower() == clean_field_name.lower():
                current_count += 1
                
                # 원하는 순번의 인스턴스 발견
                if current_count == expected_count:
                    logger.info(f"그룹 내 텍스트 일치 발견 (깊이: {depth}, 인스턴스 #{expected_count}): '{shape_text[:30]}...' -> '{content[:30]}...'")
                    change_text_to(shape_in_group, str(content))
                    return True
    
    # 여러 인스턴스 중 하나를 못 찾았어도 다른 인스턴스가 있는지 확인
    if current_count > 0:
        logger.info(f"경고: 그룹 내에서 '{clean_field_name}' 텍스트를 가진 요소를 {current_count}개 찾았지만, 요청한 {expected_count}번째 인스턴스는 찾지 못했습니다.")
    
    return False

def update_slide(slide, schema) -> None:
    """Update slide content using LLM result."""
    logger.info(f"update slide is satrting {schema['fields']} ")
    
    # 필드 정보에 그룹 컨텍스트가 포함되어 있는지 확인하는 플래그
    has_group_context = any("group_context" in field_info for field_info in schema['fields'].values() if isinstance(field_info, dict))
    
    # 메타데이터 전체를 함수에서 참조할 수 있도록 저장
    process_regular_shapes.meta_json = schema
    update_table_cell.meta_json = schema
    
    for field_name in schema['fields']:
        logger.info(f"========field name ::{field_name}================================")
        logger.info(field_name)
        logger.info(type(field_name))

        field_info = schema['fields'].get(field_name, {})
        content = field_info.get('role', "")
        field_type = field_info.get('type', "")
        group_context = field_info.get('group_context', None)
        text_count = field_info.get('text_count', 1)  # 메타데이터에서 카운트 정보 가져오기
        
        logger.info(f"=======content : ==========={content}=====================")
        if isinstance(content, list):
            content = "\n".join(content)
        
        # 필드 이름 정규화 (공백 제거, 소문자로 변환)
        normalized_field_name = ''.join(field_name.lower().split())
        
        # 태그/라벨 요소 확인 및 처리
        if is_tag_identifier(field_name) or is_tag_identifier(content):
            logger.info(f"태그/라벨 요소 처리 시도: {field_name}")
            if process_tag_element(slide, field_name, content, normalized_field_name):
                continue
        
        # 표 셀 정보 확인 (있는 경우)
        table_info = field_info.get('table_info', None)
        if table_info:
            if update_table_cell(slide, field_name, content, table_info, normalized_field_name):
                continue
        
        # 그룹 컨텍스트 정보가 있을 경우 해당 정보를 활용
        if has_group_context and group_context:
            logger.info(f"그룹 컨텍스트 정보 발견: {group_context}")
            # 그룹 요소 정보가 있으면 먼저 그룹 처리
            if process_GROUPs(slide, field_name, content, normalized_field_name):
                continue
        
        # 일반 도형들 처리 (메타데이터의 text_count 정보 활용)
        if process_regular_shapes(slide, field_name, content, normalized_field_name, text_count):
            continue
            
        # 그룹 요소 처리 (그룹 컨텍스트 정보가 없는 경우에도 시도)
        process_GROUPs(slide, field_name, content, normalized_field_name)

def load_json(file_path): 
    with open(file_path, 'r', encoding='utf-8') as f:
        schema = json.load(f)
        return schema 

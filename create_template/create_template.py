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


# Configure logging
logging.basicConfig(
    filename="logging_test.log", 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def change_text_to(shape, new_text: str) -> None:
    """
    Change text in a shape while preserving font style.
    
    Args:
        shape: The PowerPoint shape to modify
        new_text: The new text to set
    """
    try:
        text_frame = shape.text_frame
        if not text_frame:
            logger.warning(f"Shape has no text frame")
            return
            
        if not text_frame.paragraphs:
            logger.warning(f"Shape has no paragraphs")
            return
        
        # 폰트 속성을 저장하기 위한 딕셔너리
        font_props = {}
        font_found = False
        
        # 텍스트가 있는 런에서 폰트 속성 찾기
        for paragraph in text_frame.paragraphs:
            if font_found:
                break
                
            for run in paragraph.runs:
                # 텍스트가 있는 런만 확인
                if run.text and run.text.strip():
                    logger.debug(f"Found run with text: '{run.text[:20]}...'")
                    font_props['name'] = run.font.name
                    font_props['size'] = run.font.size
                    font_props['bold'] = run.font.bold
                    font_props['italic'] = run.font.italic
                    
                    # 색상 정보 저장
                    if hasattr(run.font, 'color') and run.font.color:
                        print(f"\n===== 색상 정보 디버깅 =====")
                        print(f"텍스트: '{run.text}'")
                        print(f"색상 객체: {run.font.color}")
                        print("=============================\n")
                        
                        if hasattr(run.font.color, 'rgb') and run.font.color.rgb:
                            # RGBColor 객체가 아닌 실제 정수값 저장
                            if isinstance(run.font.color.rgb, RGBColor):
                                # Python-pptx에서 RGBColor는 r, g, b 속성을 직접 노출하지 않음
                                # 대신 RGB 값 자체를 사용
                                font_props['color_rgb'] = run.font.color.rgb
                                print(f"RGBColor 객체 저장: {run.font.color.rgb}")
                            else:
                                font_props['color_rgb'] = run.font.color.rgb
                            print(f"RGB 값: {font_props['color_rgb']} (16진수: {hex(font_props['color_rgb'] if isinstance(font_props['color_rgb'], int) else 0)})")
                            r = (font_props['color_rgb'] >> 16) & 255 if isinstance(font_props['color_rgb'], int) else 0
                            g = (font_props['color_rgb'] >> 8) & 255 if isinstance(font_props['color_rgb'], int) else 0
                            b = font_props['color_rgb'] & 255 if isinstance(font_props['color_rgb'], int) else 0
                            print(f"RGB 색상: R:{r}, G:{g}, B:{b}")
                            logger.debug(f"Found color RGB: {font_props['color_rgb']}")
                        elif hasattr(run.font.color, 'theme_color') and run.font.color.theme_color:
                            font_props['theme_color'] = run.font.color.theme_color
                            print(f"테마 색상: {run.font.color.theme_color}")
                            logger.debug(f"Found theme color: {run.font.color.theme_color}")
                        elif hasattr(run.font.color, 'type'):
                            color_type = run.font.color.type
                            print(f"색상 타입: {color_type}")
                            
                            # PRESET 색상 타입 처리
                            if color_type == 102:  # PRESET (102)
                                print("PRESET 색상 타입 발견")
                                font_props['color_type'] = color_type
                                font_props['preset_color'] = True
                                logger.debug(f"Found PRESET color type: {color_type}")
                            else:
                                print(f"처리되지 않은 색상 타입: {color_type}")
                                # 기본값으로 검은색 사용
                                font_props['color_rgb'] = 0x000000
                        else:
                            print("RGB, theme_color, type 모두 없음. 검은색으로 설정.")
                            font_props['color_rgb'] = 0x000000
                    
                    font_found = True
                    break
        
        # 텍스트가 있는 런을 찾지 못한 경우, 기본 폰트 정보 사용 (첫 번째 단락의 첫 번째 런)
        if not font_found:
            print("No run with text found, using default font properties")
            if text_frame.paragraphs and text_frame.paragraphs[0].runs:
                run = text_frame.paragraphs[0].runs[0]
                font_props['name'] = run.font.name
                font_props['size'] = run.font.size
                font_props['bold'] = run.font.bold
                font_props['italic'] = run.font.italic
                
                # 색상은 검은색으로 직접 설정
                font_props['color_rgb'] = 0x000000
                logger.debug("Setting font color to black (0x000000)")
            else:
                # 런이 없는 경우에도 검은색 설정
                font_props['color_rgb'] = 0x000000
                logger.debug("No runs found, setting font color to black (0x000000)")
        
        logger.debug(f"Saved font properties: {font_props}")
        
        # 텍스트 프레임 비우기
        text_frame.clear()
        
        # 새 단락 가져오기 (clear 후에 자동으로 생성됨)
        new_paragraph = text_frame.paragraphs[0]
        new_run = new_paragraph.add_run()
        new_run.text = new_text
        
        # 폰트 속성 복원
        if 'name' in font_props and font_props['name']:
            new_run.font.name = font_props['name']
        if 'size' in font_props and font_props['size']:
            new_run.font.size = font_props['size']
        if 'bold' in font_props and font_props['bold'] is not None:
            new_run.font.bold = font_props['bold']
        if 'italic' in font_props and font_props['italic'] is not None:
            new_run.font.italic = font_props['italic']
            
        # 색상 복원
        if 'color_rgb' in font_props and font_props['color_rgb']:
            try:
                # 이미 RGBColor 객체라면 직접 설정
                if isinstance(font_props['color_rgb'], RGBColor):
                    new_run.font.color.rgb = font_props['color_rgb']
                else:
                    # 정수 값이면 RGBColor로 변환
                    r = (font_props['color_rgb'] >> 16) & 255 if isinstance(font_props['color_rgb'], int) else 0
                    g = (font_props['color_rgb'] >> 8) & 255 if isinstance(font_props['color_rgb'], int) else 0
                    b = font_props['color_rgb'] & 255 if isinstance(font_props['color_rgb'], int) else 0
                    new_run.font.color.rgb = RGBColor(r, g, b)
                logger.debug(f"Restored color RGB: {font_props['color_rgb']}")
            except Exception as e:
                print(f"색상 설정 오류: {e}")
                # 오류 발생 시 검은색으로 기본 설정
                new_run.font.color.rgb = RGBColor(0, 0, 0)
                logger.warning(f"Error setting RGB color: {e}. Using black instead.")
        elif 'theme_color' in font_props and font_props['theme_color']:
            new_run.font.color.theme_color = font_props['theme_color']
            logger.debug(f"Restored theme color: {font_props['theme_color']}")
        elif 'preset_color' in font_props and font_props['preset_color']:
            # PRESET 색상 복원
            if 'color_type' in font_props and font_props['color_type'] == 102:  # PRESET
                # color.type은 세터가 없어 직접 설정할 수 없음
                # 대신 검은색 RGB 값 사용
                new_run.font.color.rgb = RGBColor(0, 0, 0)
                logger.debug(f"Cannot set PRESET color type directly, using black instead")
            else:
                # PRESET 색상을 명확히 복원할 수 없는 경우 검은색으로 설정
                new_run.font.color.rgb = RGBColor(0, 0, 0)
                logger.debug("Could not restore PRESET color, using black instead")
        else:
            # 어떤 색상 정보도 없는 경우 검은색으로 설정
            new_run.font.color.rgb = RGBColor(0, 0, 0)
            logger.debug("No color information found, using black")
        
        logger.debug(f"Text update completed for shape: '{new_text[:30]}...'")
    except Exception as e:
        logger.error(f"Error updating text in shape: {str(e)}")


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


def update_slide(slide, schema) -> None:
    """Update slide content using LLM result."""
    logger.info(f"update slide is satrting {schema['fields']} ")
    for field_name in schema['fields']:
        logger.info(f"========field name ::{field_name}================================")
        logger.info(field_name)
        logger.info(type(field_name))

        content = schema['fields'].get(field_name, "").get('role', "")
        field_type = schema['fields'].get(field_name, {}).get('type', "")
        
        logger.info(f"=======content : ==========={content}=====================")
        if isinstance(content, list):
            content = "\n".join(content)
        
        # 필드 이름 정규화 (공백 제거, 소문자로 변환)
        normalized_field_name = ''.join(field_name.lower().split())
        
        # 표 셀 정보 확인 (있는 경우)
        table_info = schema['fields'].get(field_name, {}).get('table_info', None)
        if table_info:
            update_table_cell(slide, field_name, content, table_info, normalized_field_name)
            continue
            
        # 일반 도형들 처리
        if process_regular_shapes(slide, field_name, content, normalized_field_name):
            continue
            
        # 그룹 요소 처리
        process_group_shapes(slide, field_name, content, normalized_field_name)

def update_table_cell(slide, field_name: str, content: str, table_info: dict, normalized_field_name: str) -> bool:
    """표 셀의 텍스트를 업데이트합니다."""
    row = table_info.get('row', 0)
    col = table_info.get('col', 0)
    
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            try:
                # 표 행렬이 충분히 큰지 확인
                if shape.table.rows.count >= row and shape.table.columns.count >= col:
                    # 1부터 시작하므로 인덱스는 1 빼기
                    cell = shape.table.cell(row-1, col-1)
                    cell_text = "\n".join([p.text.strip() for p in cell.text_frame.paragraphs if p.text.strip()])
                    
                    # 셀 텍스트 정규화
                    normalized_cell_text = ''.join(cell_text.lower().split())
                    
                    if cell_text and (normalized_cell_text == normalized_field_name or cell_text.lower() == field_name.lower()):
                        logger.info(f"표 셀 일치 발견 [행:{row}, 열:{col}]: '{cell_text[:30]}...' -> '{content[:30]}...'")
                        change_text_to(cell, str(content))
                        return True
            except Exception as e:
                logger.error(f"표 셀 업데이트 중 오류: {e}")
    
    return False

def process_regular_shapes(slide, field_name: str, content: str, normalized_field_name: str) -> bool:
    """일반 도형들의 텍스트를 처리합니다."""
    for shape in slide.shapes:
        # 그룹이나 표가 아닌 일반 도형만 처리
        if shape.shape_type != MSO_SHAPE_TYPE.GROUP and shape.shape_type != MSO_SHAPE_TYPE.TABLE:
            if not hasattr(shape, "text_frame") or not shape.text_frame:
                continue
                
            # generate_meta_info.py와 동일한 방식으로 텍스트 추출
            shape_text = "\n".join([p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()])
            
            if not shape_text:
                continue
                
            # 텍스트 정규화 (공백 제거, 소문자로 변환)
            normalized_shape_text = ''.join(shape_text.lower().split())
            
            # 정규화된 텍스트 비교 (대소문자 및 공백 차이 무시)
            if normalized_shape_text == normalized_field_name:
                logger.info(f"텍스트 일치 발견: '{shape_text[:30]}...' -> '{content[:30]}...'")
                change_text_to(shape, str(content))
                return True
            # 전체 텍스트 비교가 실패하면 원본 텍스트 비교 시도 (대소문자만 무시)
            elif shape_text.lower() == field_name.lower():
                logger.info(f"원본 텍스트 일치 발견: '{shape_text[:30]}...' -> '{content[:30]}...'")
                change_text_to(shape, str(content))
                return True
    
    return False

def process_group_shapes(slide, field_name: str, content: str, normalized_field_name: str) -> bool:
    """그룹 요소 내부의 도형들을 처리합니다."""
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                for shape_in_group in shape.shapes:
                    if not hasattr(shape_in_group, "text_frame") or not shape_in_group.text_frame:
                        continue
                    
                    # 그룹 내 도형 텍스트 추출
                    shape_text = "\n".join([p.text.strip() for p in shape_in_group.text_frame.paragraphs if p.text.strip()])
                    
                    if not shape_text:
                        continue
                    
                    # 텍스트 정규화
                    normalized_shape_text = ''.join(shape_text.lower().split())
                    
                    # 텍스트 비교
                    if normalized_shape_text == normalized_field_name:
                        logger.info(f"그룹 내 텍스트 일치 발견: '{shape_text[:30]}...' -> '{content[:30]}...'")
                        change_text_to(shape_in_group, str(content))
                        return True
                    elif shape_text.lower() == field_name.lower():
                        logger.info(f"그룹 내 원본 텍스트 일치 발견: '{shape_text[:30]}...' -> '{content[:30]}...'")
                        change_text_to(shape_in_group, str(content))
                        return True
            except Exception as e:
                logger.error(f"그룹 요소 처리 중 오류: {e}")
    
    return False


def load_json(file_path): 
    with open(file_path, 'r', encoding='utf-8') as f:
        schema = json.load(f)
        return schema 


file_name='1'
input_dir ="./input"
output_dir ="./output" 
schema_path =  f"{output_dir}/meta_info_3.json"
template_schema = load_json(schema_path)
logger.info(template_schema)
original_path = f"{input_dir}/{file_name}.pptx"
prs = load_presentation(original_path)
if not prs or not prs.slides:
    raise ValueError("Invalid or empty presentation template")

# 모든 슬라이드 처리
for slide_idx, slide in enumerate(prs.slides):
    logger.info(f"슬라이드 {slide_idx+1} 처리 중...")
    update_slide(slide, template_schema)
    
# Create output directory and save
os.makedirs(output_dir, exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
pptx_filename = f"{file_name}_{timestamp}.pptx"
pptx_path = os.path.join(output_dir, pptx_filename)
prs.save(pptx_path)
logger.info(f"Presentation saved as: {pptx_path}")
print(f"모든 슬라이드 처리 완료. 파일 저장됨: {pptx_path}")

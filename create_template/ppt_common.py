import re
import logging
from typing import Dict, List, Optional, Union, Any, Tuple

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

# 로깅 설정
logging.basicConfig(
    filename="ppt_process.log", 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

#################################################
# 공통 유틸리티 함수
#################################################

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
                        # print(f"\n===== 색상 정보 디버깅 =====")
                        # print(f"텍스트: '{run.text}'")
                        # print(f"색상 객체: {run.font.color}")
                        # print("=============================\n")
                        
                        if hasattr(run.font.color, 'rgb') and run.font.color.rgb:
                            # RGBColor 객체가 아닌 실제 정수값 저장
                            if isinstance(run.font.color.rgb, RGBColor):
                                # Python-pptx에서 RGBColor는 r, g, b 속성을 직접 노출하지 않음
                                # 대신 RGB 값 자체를 사용
                                font_props['color_rgb'] = run.font.color.rgb
                                # print(f"RGBColor 객체 저장: {run.font.color.rgb}")
                            else:
                                font_props['color_rgb'] = run.font.color.rgb
                            # print(f"RGB 값: {font_props['color_rgb']} (16진수: {hex(font_props['color_rgb'] if isinstance(font_props['color_rgb'], int) else 0)})")
                            r = (font_props['color_rgb'] >> 16) & 255 if isinstance(font_props['color_rgb'], int) else 0
                            g = (font_props['color_rgb'] >> 8) & 255 if isinstance(font_props['color_rgb'], int) else 0
                            b = font_props['color_rgb'] & 255 if isinstance(font_props['color_rgb'], int) else 0
                            # print(f"RGB 색상: R:{r}, G:{g}, B:{b}")
                            logger.debug(f"Found color RGB: {font_props['color_rgb']}")
                        elif hasattr(run.font.color, 'theme_color') and run.font.color.theme_color:
                            font_props['theme_color'] = run.font.color.theme_color
                            # print(f"테마 색상: {run.font.color.theme_color}")
                            logger.debug(f"Found theme color: {run.font.color.theme_color}")
                        elif hasattr(run.font.color, 'type'):
                            color_type = run.font.color.type
                            # print(f"색상 타입: {color_type}")
                            
                            # PRESET 색상 타입 처리
                            if color_type == 102:  # PRESET (102)
                                # print("PRESET 색상 타입 발견")
                                font_props['color_type'] = color_type
                                font_props['preset_color'] = True
                                logger.debug(f"Found PRESET color type: {color_type}")
                            else:
                                # print(f"처리되지 않은 색상 타입: {color_type}")
                                # 기본값으로 검은색 사용
                                font_props['color_rgb'] = 0x000000
                        else:
                            # print("RGB, theme_color, type 모두 없음. 검은색으로 설정.")
                            font_props['color_rgb'] = 0x000000
                    
                    font_found = True
                    break
        
        # 텍스트가 있는 런을 찾지 못한 경우, 기본 폰트 정보 사용 (첫 번째 단락의 첫 번째 런)
        if not font_found:
            # print("No run with text found, using default font properties")
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


def get_shape_position(shape, slide_width, slide_height):
    """슬라이드 내의 도형 위치를 백분율로 반환합니다."""
    return {
        "left_percent": round((shape.left / slide_width) * 100, 2),
        "top_percent": round((shape.top / slide_height) * 100, 2),
        "width_percent": round((shape.width / slide_width) * 100, 2),
        "height_percent": round((shape.height / slide_height) * 100, 2)
    }


def get_type_info(shape):
    """도형의 유형 이름을 반환합니다."""
    return MSO_SHAPE_TYPE(shape.shape_type).name


def make_element_id(slide_number, pos, type_name):
    """요소의 고유 ID를 생성합니다."""
    return f"element_slide{slide_number}_l{int(pos['left_percent'])}_t{int(pos['top_percent'])}_w{int(pos['width_percent'])}_h{int(pos['height_percent'])}_type{type_name}"


#################################################
# 태그/라벨 관련 함수
#################################################

def is_tag_or_label(shape_text, pos, type_name):
    """
    텍스트가 태그나 라벨인지 판단합니다.
    
    판단 기준:
    1. 짧은 텍스트 (10자 이하)
    2. 작은 크기 (너비가 전체 슬라이드의 20% 이하)
    3. 특정 키워드 포함 (tag, label, cic, 법인 등)
    
    Args:
        shape_text: 텍스트 내용
        pos: 위치 정보 딕셔너리
        type_name: 도형 유형
        
    Returns:
        bool: 태그/라벨 여부
        str: 태그/라벨 유형 (tag, label 등)
    """
    # 텍스트 길이 확인
    is_short_text = len(shape_text) <= 10
    
    # 크기 확인
    is_small_width = pos["width_percent"] <= 20
    is_small_height = pos["height_percent"] <= 10
    
    # 특정 키워드 확인
    keywords = ["tag", "label", "cic", "법인", "id", "번호", "code", "타입", "type"]
    lowercase_text = shape_text.lower()
    has_keyword = any(keyword in lowercase_text for keyword in keywords)
    
    # 특수 패턴 확인 (예: tag_1, label_2)
    pattern_match = re.match(r'(tag|label|cic|id|code)[\s_\-]?\d*', lowercase_text)
    
    # 태그/라벨 유형 결정
    tag_type = None
    if "tag" in lowercase_text or pattern_match and "tag" in pattern_match.group():
        tag_type = "tag"
    elif "label" in lowercase_text or pattern_match and "label" in pattern_match.group():
        tag_type = "label"
    elif "cic" in lowercase_text or "법인" in shape_text:
        tag_type = "cic_label"
    elif is_short_text and (is_small_width or is_small_height):
        tag_type = "ui_element"  # 기본값
    
    # 조건 종합
    is_tag = (is_short_text and (is_small_width or is_small_height)) and (has_keyword or pattern_match)
    
    # 직접적인 패턴 매치나 키워드가 없어도 매우 작고 짧은 텍스트는 UI 요소로 간주
    if not is_tag and is_short_text and is_small_width and is_small_height:
        is_tag = True
        tag_type = "ui_element"
    
    return is_tag, tag_type


def is_tag_identifier(field_name):
    """
    필드 이름이 태그 관련 식별자인지 확인합니다.
    
    Args:
        field_name: 확인할 필드 이름
        
    Returns:
        bool: 태그 식별자 여부
    """
    # 태그 관련 키워드 확인
    tag_keywords = ["tag", "label", "cic_label", "ui_element"]
    return any(keyword in field_name.lower() for keyword in tag_keywords)


def find_tag_element(slide, normalized_field_name):
    """
    슬라이드에서 태그/라벨 요소를 찾습니다.
    태그는 크기가 작고 텍스트가 짧은 특성이 있어 부분 일치로도 검색합니다.
    
    Args:
        slide: 검색할 슬라이드
        normalized_field_name: 정규화된 필드 이름
        
    Returns:
        shape: 찾은 도형 객체 또는 None
    """
    # 태그/라벨 요소의 텍스트 길이 제한 (일반적으로 태그는 짧음)
    MAX_TAG_LENGTH = 15
    
    # 정확히 일치하는 요소 먼저 검색
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or not shape.text_frame:
            continue
            
        shape_text = "\n".join([p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()])
        if not shape_text or len(shape_text) > MAX_TAG_LENGTH:
            continue
            
        normalized_shape_text = ''.join(shape_text.lower().split())
        if normalized_shape_text == normalized_field_name:
            return shape
    
    # 정확히 일치하지 않으면 부분 일치 검색 (태그의 경우 tag_1, tag_2와 같은 형태일 수 있음)
    match_pattern = re.compile(r'(tag|label|cic|ui)[_\-\s]?\d*')
    
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or not shape.text_frame:
            continue
            
        shape_text = "\n".join([p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()])
        if not shape_text or len(shape_text) > MAX_TAG_LENGTH:
            continue
            
        normalized_shape_text = ''.join(shape_text.lower().split())
        
        # 크기가 작은지 확인 (태그/라벨은 일반적으로 작음)
        if hasattr(shape, 'width') and hasattr(shape, 'height'):
            slide_width = slide.part.slide_width
            slide_height = slide.part.slide_height
            
            width_percent = (shape.width / slide_width) * 100
            height_percent = (shape.height / slide_height) * 100
            
            is_small = width_percent <= 20 and height_percent <= 10
            
            # 작은 요소이고 패턴이 일치하면 반환
            if is_small and (match_pattern.match(normalized_shape_text) or 
                             match_pattern.match(normalized_field_name)):
                return shape
                
            # 작은 요소이고 두 텍스트가 비슷하면 반환 (편집 거리 활용)
            if is_small and len(normalized_shape_text) <= 10 and len(normalized_field_name) <= 10:
                # 레벤슈타인 거리 계산 (간단 구현)
                def levenshtein_distance(s1, s2):
                    if len(s1) < len(s2):
                        return levenshtein_distance(s2, s1)
                    if len(s2) == 0:
                        return len(s1)
                    previous_row = range(len(s2) + 1)
                    for i, c1 in enumerate(s1):
                        current_row = [i + 1]
                        for j, c2 in enumerate(s2):
                            insertions = previous_row[j + 1] + 1
                            deletions = current_row[j] + 1
                            substitutions = previous_row[j] + (c1 != c2)
                            current_row.append(min(insertions, deletions, substitutions))
                        previous_row = current_row
                    return previous_row[-1]
                
                # 짧은 텍스트에 대해 편집 거리가 작으면 유사하다고 판단
                distance = levenshtein_distance(normalized_shape_text, normalized_field_name)
                if distance <= 3:  # 편집 거리 임계값
                    return shape
    
    return None


#################################################
# 도형 검색 및 텍스트 처리 관련 함수
#################################################

def find_shape_by_text_with_count(slide, field_name, normalized_field_name, expected_count=1):
    """
    텍스트와 카운트를 기준으로 도형을 찾습니다.
    같은 텍스트의 여러 인스턴스 중 원하는 순번(count)의 것을 반환합니다.
    
    Args:
        slide: 검색할 슬라이드
        field_name: 원본 필드 이름
        normalized_field_name: 정규화된 필드 이름
        expected_count: 찾으려는 인스턴스의 카운트 번호 (1부터 시작)
        
    Returns:
        tuple: (찾은 도형, 현재 카운트)
    """
    current_count = 0
    
    # 모든 도형 조사
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or not shape.text_frame:
            continue
            
        shape_text = "\n".join([p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()])
        if not shape_text:
            continue
            
        # 정규화된 텍스트로 비교
        normalized_shape_text = ''.join(shape_text.lower().split())
        
        # 텍스트가 일치하면 카운터 증가
        if normalized_shape_text == normalized_field_name or shape_text.lower() == field_name.lower():
            current_count += 1
            
            # 원하는 순번의 인스턴스 발견
            if current_count == expected_count:
                return shape, current_count
                
    # 찾지 못한 경우
    return None, current_count


def extract_count_from_field_name(field_name):
    """
    필드 이름에서 카운트 정보 추출.
    
    Args:
        field_name: 필드 이름 (예: "title_1", "content_2")
        
    Returns:
        tuple: (수정된 필드 이름, 추출된 카운트)
    """
    count_match = re.search(r'_(\d+)$', field_name)
    expected_count = 1
    
    # 이름 끝에 _숫자 형식이 있으면 expected_count 설정
    if count_match:
        try:
            expected_count = int(count_match.group(1))
            # 매칭된 숫자 부분을 제외한 필드 이름 사용
            clean_field_name = field_name[:count_match.start()]
            clean_normalized_field_name = ''.join(clean_field_name.lower().split())
            return clean_field_name, clean_normalized_field_name, expected_count
        except (ValueError, IndexError):
            pass
    
    # 숫자 형식이 없으면 원본 필드 이름 그대로 사용
    return field_name, ''.join(field_name.lower().split()), expected_count 

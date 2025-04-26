import json
import os

def read_ppt_analysis(analysis_path):
    """PPT 분석 결과를 읽습니다."""
    with open(analysis_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def generate_prompt_for_object(object_info):
    """객체의 설명을 생성하기 위한 프롬프트를 생성합니다."""
    prompt = f"""
다음은 PPT 슬라이드의 객체 정보입니다:
- 텍스트 내용: {object_info['identifier'].get('text_content', '없음')}
- 위치: 상단 {object_info['identifier']['position']['relative_position']['top_percent']}%, 
        왼쪽 {object_info['identifier']['position']['relative_position']['left_percent']}%
- 크기: 너비 {object_info['identifier']['position']['relative_position']['width_percent']}%, 
        높이 {object_info['identifier']['position']['relative_position']['height_percent']}%
- 텍스트 길이: {object_info['identifier'].get('text_length', 0)}자
- 텍스트 스타일: {object_info['identifier'].get('text_style', {})}

이 객체가 PPT에서 어떤 역할을 하는지 설명해주세요.
다음과 같은 요소들을 고려해주세요:
1. 위치 (상단/중앙/하단)
2. 크기 (큰/중간/작은)
3. 텍스트 내용의 성격 (제목/부제목/본문/목차 등)
4. 텍스트 길이
5. 텍스트 스타일 (글꼴 크기, 굵기 등)

설명은 "~을 위한 객체" 형식으로 작성해주세요.
"""
    return prompt

def add_descriptions_to_analysis(analysis_data):
    """분석 데이터에 설명을 추가합니다."""
    for slide in analysis_data['slides']:
        for shape in slide['shapes']:
            if 'identifier' in shape and 'text_content' in shape['identifier']:
                prompt = generate_prompt_for_object(shape)
                # TODO: 실제로는 여기서 AI 모델을 호출하여 설명을 생성
                # 현재는 임시로 기본 설명을 사용
                shape['identifier']['description'] = generate_basic_description(shape)
    return analysis_data

def generate_basic_description(shape):
    """기본적인 설명을 생성합니다. (임시 함수)"""
    position = shape['identifier']['position']['relative_position']
    text_content = shape['identifier']['text_content']
    
    # 위치 기반 설명
    position_desc = ""
    if position['top_percent'] < 20:
        position_desc = "상단에 위치한 "
    elif position['top_percent'] > 80:
        position_desc = "하단에 위치한 "
    
    # 크기 기반 설명
    size_desc = ""
    if position['width_percent'] > 70:
        size_desc = "큰 "
    elif position['width_percent'] < 30:
        size_desc = "작은 "
    
    # 텍스트 내용 기반 설명
    content_desc = ""
    if "제목" in text_content or "Title" in text_content:
        content_desc = "제목"
    elif "부제목" in text_content or "Subtitle" in text_content:
        content_desc = "부제목"
    elif "목차" in text_content or "Contents" in text_content:
        content_desc = "목차"
    elif "참고" in text_content or "Reference" in text_content:
        content_desc = "참고 문헌"
    elif "결론" in text_content or "Conclusion" in text_content:
        content_desc = "결론"
    elif "소개" in text_content or "Introduction" in text_content:
        content_desc = "소개"
    else:
        content_desc = "본문"
    
    return f"{position_desc}{size_desc}{content_desc}을 위한 객체"

def save_schema(schema, output_path):
    """스키마를 JSON 파일로 저장합니다."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(schema, f, ensure_ascii=False, indent=2)

def main():
    analysis_path = "ppt_analysis.json"
    output_path = "ppt_schema_with_descriptions.json"
    
    try:
        # 1. PPT 분석 결과 읽기
        analysis_data = read_ppt_analysis(analysis_path)
        
        # 2. 설명 추가
        schema_with_descriptions = add_descriptions_to_analysis(analysis_data)
        
        # 3. 결과 저장
        save_schema(schema_with_descriptions, output_path)
        print(f"설명이 추가된 스키마가 {output_path}에 저장되었습니다.")
        
        # 4. Git에 푸시
        os.system("git add .")
        os.system("git commit -m 'Add PPT schema with descriptions'")
        os.system("git push")
        print("변경사항이 Git에 푸시되었습니다.")
        
    except Exception as e:
        print(f"에러 발생: {str(e)}")

if __name__ == "__main__":
    main() 
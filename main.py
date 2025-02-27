import os
from openai import OpenAI
from dotenv import load_dotenv
import docx
import pandas as pd
import glob
from tqdm import tqdm

# 加载环境变量
load_dotenv()

# 初始化OpenAI客户端
client = OpenAI(
    api_key=os.environ.get("ARK_API_KEY"),
    base_url="https://ark.cn-beijing.volces.com/api/v3",
)

# 用于存储所有生成的题目
all_questions = []

# 处理单个Word文档
def process_doc_file(doc_path):
    print(f"处理文件: {doc_path}")
    doc = docx.Document(doc_path)
    # 提取文档内容
    content = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
    
    # 文档基础信息
    filename = os.path.basename(doc_path)
    category = filename.split('.')[0]  # 使用文件名作为分类
    
    # 分批次生成题目，每次5道
    num_batches = 1 # 25道题/5=5批
    questions_per_batch = 25
    
    for batch in range(num_batches):
        # 计算当前批次应生成的单选题和多选题数量
        if batch < 4:  # 前4批各5道题，共20道
            single_choice = 20  # 每批4道单选
            multiple_choice = 5  # 每批1道多选
        else:  # 最后1批5道题
            single_choice = 20  # 3道单选
            multiple_choice = 5  # 2道多选
        
        # 生成提示
        prompt = f"""
        你是一个网络安全培训机构的专业出题老师。请根据以下内容，创建{single_choice}道单选题和{multiple_choice}道多选题，总共{questions_per_batch}道题。
        
        文档内容:
        {content[:4000]}  # 限制内容长度以避免超过token限制
        
        要求:
        1. 题目必须基于上述内容, 题目不能有xxx的比喻/类比这种类似的问题。
        2. 单选题只有一个正确答案,多选题有2-4个正确答案
        3. 每道题必须包括题目、选项A-D(最多到F)、正确答案和解析
        4. 以JSON格式返回,每道题包括以下字段:
          - type: "单选题"或"多选题"
          - question: 题目内容
          - options: 包含选项A-D(或更多)的对象
          - answer: 正确答案，如"A"或"ABCD"
          - explanation: 试题解析
        5. 题目号从{batch * questions_per_batch + 1}开始
        
        返回格式示例:
        [
          {{
            "number": 1,
            "category": "{category}",
            "type": "单选题",
            "question": "题目内容",
            "options": {{
              "A": "选项A内容",
              "B": "选项B内容",
              "C": "选项C内容",
              "D": "选项D内容"
            }},
            "answer": "A",
            "explanation": "解析内容"
          }},
          ...
        ]
        
        请确保JSON格式正确，可以被Python直接解析。每道多选题的正确答案必须在2-4个之间，不能所有选项都是正确答案。
        """
        
        try:
            # 调用API生成题目
            response = client.chat.completions.create(
                model="deepseek-v3-241226",
                messages=[
                    {"role": "system", "content": "你是一个专业的网络安全培训题目生成器"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
            )
            
            # 解析响应并添加到结果列表
            result = response.choices[0].message.content
            
            # 提取JSON部分
            import json
            import re
            
            # 找到JSON数组部分
            json_match = re.search(r'\[\s*\{.*\}\s*\]', result, re.DOTALL)
            if json_match:
                json_str = json_match.group(0)
                questions = json.loads(json_str)
                
                # 添加到全局列表
                all_questions.extend(questions)
                print(f"已完成第{batch+1}批，生成了{len(questions)}道题目")
            else:
                print(f"无法解析JSON格式，跳过第{batch+1}批")
                print("API返回内容:", result)
        
        except Exception as e:
            print(f"处理批次{batch+1}时出错: {str(e)}")
            
    return len(all_questions)

# 将题目转换为Excel格式
def convert_to_excel_format(questions):
    excel_data = []
    for q in questions:
        row = {
            "必填-题号": q["number"],
            "必填-分类": q["category"],
            "必填-题型": q["type"],
            "必填-题目": q["question"],
            "必填-标准答案": q["answer"],
            "答案A": q["options"].get("A", ""),
            "答案B": q["options"].get("B", ""),
            "答案C": q["options"].get("C", ""),
            "答案D": q["options"].get("D", ""),
            "答案E": q["options"].get("E", ""),
            "答案F": q["options"].get("F", ""),
            "试题解析": q["explanation"]
        }
        excel_data.append(row)
    
    return excel_data

# 主函数
def main():
    # 获取目录下所有doc文件
    doc_files = glob.glob("docs/*.doc*")
    
    if not doc_files:
        print("当前目录下未找到任何Word文档")
        return
    
    print(f"找到{len(doc_files)}个Word文档文件")
    
    # 处理每个文件
    for doc_file in tqdm(doc_files):
        try:
            count = process_doc_file(doc_file)
            print(f"从 {doc_file} 生成了 {count} 道题目")
        except Exception as e:
            print(f"处理文件 {doc_file} 时出错: {str(e)}")
    
    # 转换为Excel格式并保存
    if all_questions:
        excel_data = convert_to_excel_format(all_questions)
        df = pd.DataFrame(excel_data)
        
        output_file = "security_questions.xlsx"
        df.to_excel(output_file, index=False)
        print(f"所有题目已保存到 {output_file}")
    else:
        print("未生成任何题目")

if __name__ == "__main__":
    main()

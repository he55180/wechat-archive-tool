import re
import sys

def preprocess_markdown(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 正则表达式：匹配行首的 "数字. "，替换为 "数字\. "
    # 这样 Pandoc 就不会将其解析为 List Paragraph，而是普通 Text Paragraph
    # 脚本 format_expert.py 能够识别 "1." 开头的普通段落为三级标题
    
    # Pattern: Line start, optional whitespace, one or more digits, dot, space
    # Replacement: The matched digits, escaped dot, space
    # Note: We group the whitespace/digits to preserve indentation if any (though usually top level)
    
    # Using a lookahead to ensure we only target the specific structure
    # \g<1> refers to the whitespace, \g<2> to the digits
    updated_content = re.sub(r'^(\s*)(\d+)\.(\s+)', r'\1\2. ', content, flags=re.MULTILINE)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(updated_content)
    
    print(f"Processed {input_path} -> {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python preprocess_md.py input.md output.md")
        sys.exit(1)
    
    preprocess_markdown(sys.argv[1], sys.argv[2])

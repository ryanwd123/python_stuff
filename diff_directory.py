#%%
import os
# os.environ['PYWEBVIEW_GUI'] = 'edgechromium'
from pathlib import Path
import re
from turtle import ht
import webview
import diff_match_patch as dmp_module
import webview.window

def count_diffs(diffs):
    insertions = sum(1 for (op, _) in diffs if op == dmp_module.diff_match_patch.DIFF_INSERT)
    deletions = sum(1 for (op, _) in diffs if op == dmp_module.diff_match_patch.DIFF_DELETE)
    changes = sum(1 for (op, _) in diffs if op == dmp_module.diff_match_patch.DIFF_EQUAL)
    return insertions + deletions + changes



def get_html_for_files(file1:Path, file2:Path):
    dmp = dmp_module.diff_match_patch()
    text1 = file1.read_text()
    text2 = file2.read_text()
    diffs = dmp.diff_main(text1, text2)
    # dmp.diff_cleanupSemantic(diffs)
    html = dmp.diff_prettyHtml(diffs)
    count_of_diffs = count_diffs(diffs)
    print(f'{file1.name} count_of_diffs: {count_of_diffs}')
    if count_of_diffs == 1:
        return ''
    return html

def compare_folders(folder1:Path, folder2:Path):
    html = ''
    files = folder1.glob('**/*')
    for file in files:
        html_for_file = ''
        if file.is_file():
            html_for_file += f'<h1>{file.relative_to(folder1)}</h1>'
            file2 = folder2 / file.relative_to(folder1)
            if file2.exists():
                diff_html_txt = get_html_for_files(file, file2)
                if len(diff_html_txt) > 0:
                    html_for_file += diff_html_txt
                    html += html_for_file

    return html


# diff_html = get_html_for_files('text1.txt', 'text2.txt')
diff_html = compare_folders(Path('folder1'), Path('folder2'))
diff_html = diff_html.replace('#e6ffe6','lightgreen').replace('&para;','').replace('    ','&nbsp;&nbsp;&nbsp;&nbsp;')


html = f"""
<html>

<head>
    <style>
        body {{
            font-family: monospace;
            font-size: 16px;
        }}
    </style>
</head>
<body>
    {diff_html}
</body>

</html>
"""


# print(html)


w = webview.create_window('Hello world', html=html, width=800, height=800, x=-900, y=100)
webview.start()

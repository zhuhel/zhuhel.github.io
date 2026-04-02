#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
物理组科研数据库更新脚本
将 Excel 表格转换为网站所需的 JSON 数据库格式

使用方法：
1. 将新的 Excel 文件命名为 "物理组科研摸底表.xlsx" 并放在同一目录
2. 运行: python update_database.py
3. 生成的 database.json 即为网站数据文件
"""

import pandas as pd
import json
import urllib.parse
import os


def generate_paper_links(title: str, journal: str) -> dict:
    """为论文生成搜索链接"""
    if not title:
        return {}
    
    encoded_title = urllib.parse.quote(title)
    
    links = {
        'googleScholar': f'https://scholar.google.com/scholar?q={encoded_title}',
        'pubmed': f'https://pubmed.ncbi.nlm.nih.gov/?term={encoded_title}',
        'cnki': f'https://kns.cnki.net/kns8/defaultresult/index?kw={urllib.parse.quote(title)}&korder=SU'
    }
    
    # 根据期刊类型判断主要搜索源
    journal_upper = journal.upper() if journal else ''
    if any(keyword in journal_upper for keyword in ['MEDICAL', 'RADIATION', 'PHYSICS', 'JOURNAL', 'ONCOLOGY']):
        links['primary'] = 'pubmed'
    else:
        links['primary'] = 'cnki'
    
    return links


def parse_papers(xl: pd.ExcelFile) -> list:
    """解析论文数据"""
    papers = []
    
    try:
        df_papers = xl.parse('文章发表情况', header=None)
    except Exception as e:
        print(f"警告: 无法读取 '文章发表情况' 工作表: {e}")
        return papers
    
    # 从第2行开始读取（第0行是标题，第1行是列名）
    for i in range(2, len(df_papers)):
        row = df_papers.iloc[i]
        
        # 跳过空行
        if pd.isna(row[0]) and pd.isna(row[2]):
            continue
            
        paper = {
            'id': f'paper_{i-1}',
            'type': 'paper',
            'firstAuthor': str(row[0]).strip() if pd.notna(row[0]) else '',
            'correspondingAuthor': str(row[1]).strip() if pd.notna(row[1]) else '',
            'title': str(row[2]).strip() if pd.notna(row[2]) else '',
            'category': str(row[3]).strip() if pd.notna(row[3]) else '',
            'year': str(row[4]).strip() if pd.notna(row[4]) else '',
            'journal': str(row[5]).strip() if pd.notna(row[5]) else '',
            'journalType': str(row[6]).strip() if pd.notna(row[6]) else '',
            'journalQ': str(row[7]).strip() if pd.notna(row[7]) else '',
            'coAuthors': str(row[8]).strip() if pd.notna(row[8]) else '',
            'funding': str(row[9]).strip() if pd.notna(row[9]) else ''
        }
        
        # 生成论文链接
        paper['links'] = generate_paper_links(paper['title'], paper['journal'])
        papers.append(paper)
    
    return papers


def parse_patents(xl: pd.ExcelFile) -> list:
    """解析专利数据"""
    patents = []
    
    try:
        df_patents = xl.parse('授权专利情况', header=None)
    except Exception as e:
        print(f"警告: 无法读取 '授权专利情况' 工作表: {e}")
        return patents
    
    # 从第2行开始读取
    for i in range(2, len(df_patents)):
        row = df_patents.iloc[i]
        
        if pd.isna(row[0]):
            continue
            
        patent = {
            'id': f'patent_{i-1}',
            'type': 'patent',
            'firstInventor': str(row[0]).strip() if pd.notna(row[0]) else '',
            'title': str(row[1]).strip() if pd.notna(row[1]) else '',
            'number': str(row[2]).strip() if pd.notna(row[2]) else '',
            'patentType': str(row[3]).strip() if pd.notna(row[3]) else '',
            'year': str(row[4]).strip() if pd.notna(row[4]) else '',
            'cooperation': str(row[5]).strip() if pd.notna(row[5]) else '',
            'otherInventors': str(row[6]).strip() if pd.notna(row[6]) else '',
            'funding': str(row[7]).strip() if pd.notna(row[7]) else ''
        }
        patents.append(patent)
    
    return patents


def parse_funds(xl: pd.ExcelFile) -> list:
    """解析基金数据"""
    funds = []
    
    try:
        df_funds = xl.parse('获得基金情况', header=None)
    except Exception as e:
        print(f"警告: 无法读取 '获得基金情况' 工作表: {e}")
        return funds
    
    # 从第2行开始读取
    for i in range(2, len(df_funds)):
        row = df_funds.iloc[i]
        
        if pd.isna(row[0]):
            continue
            
        fund = {
            'id': f'fund_{i-1}',
            'type': 'fund',
            'principal': str(row[0]).strip() if pd.notna(row[0]) else '',
            'title': str(row[1]).strip() if pd.notna(row[1]) else '',
            'fundInfo': str(row[2]).strip() if pd.notna(row[2]) else '',
            'duration': str(row[3]).strip() if pd.notna(row[3]) else '',
            'amount': str(row[4]).strip() if pd.notna(row[4]) else '',
            'participants': str(row[5]).strip() if pd.notna(row[5]) else ''
        }
        funds.append(fund)
    
    return funds


def parse_projects(xl: pd.ExcelFile) -> list:
    """解析在研课题数据"""
    projects = []
    
    try:
        df_projects = xl.parse('在研课题', header=None)
    except Exception as e:
        print(f"警告: 无法读取 '在研课题' 工作表: {e}")
        return projects
    
    # 从第2行开始读取
    for i in range(2, len(df_projects)):
        row = df_projects.iloc[i]
        
        if pd.isna(row[0]):
            continue
            
        project = {
            'id': f'project_{i-1}',
            'type': 'project',
            'person': str(row[0]).strip() if pd.notna(row[0]) else '',
            'expectedOutcome': str(row[1]).strip() if pd.notna(row[1]) else '',
            'duration': str(row[2]).strip() if pd.notna(row[2]) else '',
            'cooperation': str(row[3]).strip() if pd.notna(row[3]) else '',
            'status': str(row[4]).strip() if pd.notna(row[4]) else '',
            'content': str(row[5]).strip() if pd.notna(row[5]) else ''
        }
        projects.append(project)
    
    return projects


def extract_people(papers: list, projects: list, patents: list = None, funds: list = None) -> list:
    """从所有数据中提取人员列表"""
    all_people = set()
    
    # 从论文中提取人员
    for paper in papers:
        for field in [paper.get('firstAuthor', ''), paper.get('correspondingAuthor', ''), paper.get('coAuthors', '')]:
            if field:
                for name in field.replace('，', ',').replace('、', ',').split(','):
                    name = name.strip()
                    if name and len(name) >= 2:
                        all_people.add(name)
    
    # 从在研课题中提取人员
    for project in projects:
        if project.get('person'):
            for name in project['person'].replace('，', ',').replace('、', ',').split(','):
                name = name.strip()
                if name and len(name) >= 2:
                    all_people.add(name)
    
    # 从专利中提取人员
    if patents:
        for patent in patents:
            if patent.get('firstInventor'):
                for name in patent['firstInventor'].replace('，', ',').replace('、', ',').split(','):
                    name = name.strip()
                    if name and len(name) >= 2:
                        all_people.add(name)
            if patent.get('otherInventors'):
                for name in patent['otherInventors'].replace('，', ',').replace('、', ',').split(','):
                    name = name.strip()
                    if name and len(name) >= 2:
                        all_people.add(name)
    
    # 从基金中提取人员
    if funds:
        for fund in funds:
            if fund.get('principal'):
                for name in fund['principal'].replace('，', ',').replace('、', ',').split(','):
                    name = name.strip()
                    if name and len(name) >= 2:
                        all_people.add(name)
            if fund.get('participants'):
                for name in fund['participants'].replace('，', ',').replace('、', ',').split(','):
                    name = name.strip()
                    if name and len(name) >= 2:
                        all_people.add(name)
    
    return sorted(list(all_people))


def main():
    """主函数"""
    excel_file = '物理组科研摸底表.xlsx'
    
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        print(f"错误: 找不到文件 '{excel_file}'")
        print("请将 Excel 文件命名为 '物理组科研摸底表.xlsx' 并放在同一目录下")
        return
    
    print(f"正在读取: {excel_file}")
    
    try:
        xl = pd.ExcelFile(excel_file)
        print(f"发现工作表: {', '.join(xl.sheet_names)}")
    except Exception as e:
        print(f"错误: 无法读取 Excel 文件: {e}")
        return
    
    # 解析各类数据
    print("\n正在解析数据...")
    
    papers = parse_papers(xl)
    print(f"  ✓ 论文: {len(papers)} 篇")
    
    patents = parse_patents(xl)
    print(f"  ✓ 专利: {len(patents)} 项")
    
    funds = parse_funds(xl)
    print(f"  ✓ 基金: {len(funds)} 项")
    
    projects = parse_projects(xl)
    print(f"  ✓ 在研课题: {len(projects)} 个")
    
    # 提取人员
    people = extract_people(papers, projects, patents, funds)
    print(f"  ✓ 人员: {len(people)} 人")
    
    # 构建数据库
    database = {
        'papers': papers,
        'patents': patents,
        'funds': funds,
        'projects': projects,
        'people': people
    }
    
    # 保存 JSON
    output_file = 'database.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(database, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 数据库已生成: {output_file}")
    print(f"\n数据统计:")
    print(f"  - 论文: {len(papers)} 篇")
    print(f"  - 专利: {len(patents)} 项")
    print(f"  - 基金: {len(funds)} 项")
    print(f"  - 在研课题: {len(projects)} 个")
    print(f"  - 人员: {len(people)} 人")
    print(f"\n下一步:")
    print(f"  将 {output_file} 上传到网站服务器，替换旧文件即可")


if __name__ == '__main__':
    main()

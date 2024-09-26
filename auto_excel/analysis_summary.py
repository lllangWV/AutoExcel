import json
import logging
import os
import sys

from dotenv import load_dotenv
from functools import partial

load_dotenv()

# from llama_index.agent import OpenAIAgent
from llama_index.core.tools import QueryEngineTool, ToolMetadata
from llama_index.llms.openai import OpenAI
from llama_index.llms.anthropic import Anthropic

from llama_index.core.llms import ChatMessage

from llama_index.experimental.query_engine import PandasQueryEngine
import pandas as pd

from auto_excel.analysis import extract_disjoint_tables

logger = logging.getLogger(__name__)

logger.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch.setFormatter(formatter)
logger.addHandler(ch)


def write_llm_query_response(query, response, context=None, output_dir=None, metadata=None):
    os.makedirs(output_dir, exist_ok=True)
    
    with open(os.path.join(output_dir, f'query_reponse.md'), 'w') as f:
        f.write(f'# Query: \n\n')
        f.write(f'{query}\n\n')
        f.write(f'# Response: \n\n')
        f.write(f'{response}\n\n')

    with open(os.path.join(output_dir, f'query.md'), 'w') as f:
        f.write(f'# Query: \n\n')
        f.write(f'{query}\n\n')

    with open(os.path.join(output_dir, f'response.md'), 'w') as f:
        f.write(f'# Response: \n\n')
        f.write(f'{response}\n\n')

    if context:
        with open(os.path.join(output_dir, f'context.md'), 'w') as f:
            f.write(f'# Context: \n\n')
            f.write(f'{context}\n\n')

    if metadata:
        with open(os.path.join(output_dir, f'metadata.json'), 'w') as f:
            json.dump(metadata, f, indent=4)
                

def main():

    #########################################################################################
    # Parameters
    file_path = 'C:/Users/lllang/Desktop/Current_Projects/Auto_Excel/data/processed/Data Analysis 9-19-2024.xlsx'
    
    llm_companys=['anthropic', 'openai']
    llm_company=llm_companys[0]
    if llm_company=='anthropic':
        print('Using Antrhopic')
        models=['claude-3-5-sonnet-20240620','claude-3-opus-20240229','claude-3-sonnet-20240229','claude-3-haiku-20240307']
        model=models[0]
        llm=Anthropic(model=model)
    elif llm_company=='openai':
        print('Using OpenAI')
        openai_models=['gpt-4o', 'gpt-4o-mini', 'o1-mini', 'o1-preview']
        model=openai_models[0]
        llm = OpenAI(model=model)

    metadata={'llm_company':llm_company, 
              'llm_model':model,
              'filepath':file_path}

    #########################################################################################
    # Script starts here
    #########################################################################################

    output_dir=os.path.join('data','output')
    os.makedirs(output_dir, exist_ok=True)

    n_directories=len(os.listdir(output_dir))
    directory_name=f'summary_report_{n_directories}'
    response_dir=os.path.join(output_dir,directory_name)
    os.makedirs(response_dir, exist_ok=True)
    #########################################################################################
    # Create the query engine
    fy_analytics_tables = extract_disjoint_tables(file_path, sheet_name='FY 24-25 Analytics')
    caseload_analysis_tables = extract_disjoint_tables(file_path, sheet_name='Caseload Analysis')
    # fy_sharepoint_tables = extract_disjoint_tables(file_path, sheet_name='FY 24 SharePoint')
    # active_assignments_tables = extract_disjoint_tables(file_path, sheet_name='Active Assignments')

    # df=fy_sharepoint_tables[0]
    # df= pd.read_excel(file_path, sheet_name='FY 24 SharePoint')
    # print(df.shape)
    tables=fy_analytics_tables
    tables.extend(caseload_analysis_tables)
    # df = pd.DataFrame(tables)

    context="Context:\n\n"
    for i,df in enumerate(tables):
        context+='Table {}: \n Shape: {}\n\n'.format(i+1,df.shape)
        context+=df.to_markdown()
        context+='\n\n'
        

    #########################################################################################


    instruction_str = """\
    
    1. Read context of the tables in the Conext section.
    2. Write a summary report on the information in the tables.
    3. Write in markdown format.
    4. Write in paragraph format. Keepthe number of tables in the report limited to the most important information.
    5. PRINT ONLY THE EXPRESSION.
    """
    prompt_template = """\
    
    You are analyzing a financial analytics for a financial year. There are multiple tables with information ranging 
    total assignments by delinqunecy types, Executed Agreements by Negotiator, New Assignments by Negotiator,etc:

    You're goal is to provide a in-depth summary report.

    Follow these instructions:
    {instruction_str}

    Context:

    {context}

    Markdown Summary Report: """

    # messages = [
    #     ChatMessage(
    #         role="system", content="You are a pirate with a colorful personality"
    #     ),
    #     ChatMessage(role="user", content="Tell me a story"),
    # ]


    full_query=prompt_template.format(instruction_str=instruction_str, context=context)

    
    logger.debug(f"Query: {full_query}" )

    response = llm.complete(prompt_template.format(instruction_str=instruction_str, context=context))

    metadata['input_characters'] = len(full_query)
    metadata['output_characters'] = len(str(response))


    query = prompt_template.format(instruction_str=instruction_str, context='')
    logger.info(f"Response: {str(response)}")
    logger.info(str(response))
    write_llm_query_response(query, str(response), context=context, output_dir=response_dir, metadata=metadata)





   
if __name__ == '__main__':
    main()
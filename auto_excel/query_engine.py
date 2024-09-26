import logging
import os
import sys

from dotenv import load_dotenv

load_dotenv()

# from llama_index.agent import OpenAIAgent
from llama_index.core.tools import QueryEngineTool, ToolMetadata
from llama_index.core import PromptTemplate
from llama_index.llms.openai import OpenAI
from llama_index.llms.anthropic import Anthropic


from llama_index.experimental.query_engine import PandasQueryEngine
import pandas as pd

from auto_excel.analysis import extract_disjoint_tables
from auto_excel.utils import write_llm_query_response

logging.basicConfig(stream=sys.stdout, level=logging.INFO)
logging.getLogger().addHandler(logging.StreamHandler(stream=sys.stdout))

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
    output_dir=os.path.join('data','output','spreadsheet_chat')
    os.makedirs(output_dir, exist_ok=True)

    n_directories=len(os.listdir(output_dir))
    directory_name=f'chat_{n_directories}'
    response_dir=os.path.join(output_dir,directory_name)
    os.makedirs(response_dir, exist_ok=True)
    #########################################################################################
    # Create the query engine
    # fy_analytics_tables = extract_disjoint_tables(file_path, sheet_name='FY 24-25 Analytics')
    # caseload_analysis_tables = extract_disjoint_tables(file_path, sheet_name='Caseload Analysis')
    # fy_sharepoint_tables = extract_disjoint_tables(file_path, sheet_name='FY 24 SharePoint')
    # active_assignments_tables = extract_disjoint_tables(file_path, sheet_name='Active Assignments')

    # df=fy_sharepoint_tables[0]

    metadata['sheet_name']='FY 24 SharePoint'
    df= pd.read_excel(file_path, sheet_name=metadata['sheet_name'])
    print(df.shape)


    query_engine = PandasQueryEngine(df=df, 
                                     llm=llm,
                                     verbose=True)
    

    #########################################################################################
    # Updating built in prompts.
    new_prompt = PromptTemplate(
    """\
    You are working with a pandas dataframe in Python.
    The name of the dataframe is `df`.
    This is the result of `print(df.head())`:
    {df_str}

    The first column is the case number it's column name is `#`, it is an integer. It gives the case number of the assignment.
    The second column is the negotiator it's column name is `Negotiator`, it is a string. It gives the negotiator of the assignment.
    The third column is the Status it's column name is `Status`, it is a string. It gives the status of the assignment.
    The fourth column is the OSP Number it's column name is `OSP Number`, it is a string. It gives the OSP number of the assignment.
    The fifth column is the Date Assigned it's column name is `Date Assigned`, it is a string with the format `mm/dd/yyyy`. It gives the date the assignment was assigned.
    The sixth column is the Pricipal Investigator it's column name is `Pricipal Investigator`, it is a string. It gives the principal investigator of the assignment.
    The seventh column is the agreement type it's column name is `Agreement Type`, it is a string. It gives the agreement type of the assignment.
    The eighth column is the Date Received at OSP it's column name is `Date Received at OSP`, it is a string with the format `mm/dd/yyyy`. It gives the date the assignment was received at OSP.
    The ninth column is the Deadline Date it's column name is `Deadline Date`, it is a string with the format `mm/dd/yyyy`. It gives the deadline date of the assignment.
    The tenth column is the Sponsor it's column name is `Sponsor`, it is a string. It gives the sponsor of the assignment.
    The eleventh column is the Additional Information it's column name is `Additional Information`, it is a string. It gives the additional information of the assignment.
    The twelfth column is the FE Date it's column name is `FE Date`, it is a string with the format `mm/dd/yyyy`. It gives the date the assignment was forwarded to the FE.
    The thirteenth column is the High Priority its column name is `High Priority`, it is a string. It tells whether the assignment is high priority by "Yes".
    The fourteenth column is the New Negotiator Comments its column name is `New Negotiator Comments`, it is a string. It gives the new negotiator comments of the assignment.
    The fifteenth column is the Department it's column name is `Department`, it is a string. It gives the department of the assignment.
    The sixteenth column is the College it's column name is `College`, it is a string. It gives the college of the assignment.
    The seventeenth column is the Date Received at WVU it's column name is `Date Received at WVU`, it is a string with the format `mm/dd/yyyy`. It gives the date the assignment was received at WVU.
    The eighteenth column is the Time to Assignment it's column name is `Time to Assignment`, it is a string. It gives the number of days betweenn assignment date and date received at OSP.
    The nineteenth column is the Time to Execution it's column name is `Time to Execution`, it is a string. It gives the number of days betweenn assignment date and date forwarded to FE.
    The twentieth column is the Today's Date it's column name is `Today's Date`, it is a string with the format `mm/dd/yyyy`. It gives the date  the analysis was run.
    The twenty-first column is the Time Since Assignment it's column name is `Time Since Assignment`, it is a string. It gives the number of days betweenn today's date and date assigned.
    The twenty-second column is the Delinquency it's column name is `Delinquency`, it is a string. It gives the categorizes the type of delinquency. It can be "> 90 Days", "> 60 Days", "> 30 Days", or "< 30 Days".

    Follow these instructions:
    {instruction_str}
    Query: {query_str}

    Expression: """
    )

    # metadata['input_characters'] = len(full_query)

    query_engine.update_prompts({"pandas_prompt": new_prompt})

    # This is the instruction string (that you can customize by passing in instruction_str on initialization)
    instruction_str = """\
    1. Convert the query to executable Python code using Pandas.
    2. The final line of code should be a Python expression that can be called with the `eval()` function.
    3. The code should represent a solution to the query.
    4. PRINT ONLY THE EXPRESSION.
    5. Do not quote the expression.
    """


    #########################################################################################
    # Example 1:

    # query = " I want the total assignmnets by delinquency type."
    # response = query_engine.query(query)

    # write_llm_query_response(query, str(response), output_dir=response_dir, metadata=metadata)


    #########################################################################################
    # Example 2:

    query = " I want the know the total assignmnets grouped by delinquency type and negotiator."
    response = query_engine.query(query)

    write_llm_query_response(query, str(response), output_dir=response_dir, metadata=metadata)



    #########################################################################################
    # Example 2:

    # query = " I want the number of Executed Agreements by Negotiator for 2204 by month. "
    # response = query_engine.query(query)


    # print(str(response))

    # # Create a tool that allows the agent to query the DataFrame
    # pandas_tool = QueryEngineTool.from_query_engine(
    #     query_engine=pandas_index.as_query_engine(),
    #     metadata=ToolMetadata(
    #         name='pandas_index',
    #         description='Useful for answering questions about the CSV data.'
    #     )
    # )
    # Create the agent and provide the pandas tool
    # agent = OpenAIAgent.from_tools([pandas_tool], verbose=True)

    # # Your query
    # question = "What is the high priority assignment analysis by negotiator?"

    # # Get the response from the agent
    # response = agent.chat(question)

    # # Print the response
    # print(response)
if __name__ == '__main__':
    main()
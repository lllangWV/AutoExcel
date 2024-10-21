import json
import logging
import os
import sys

logger = logging.getLogger(__name__)

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
                

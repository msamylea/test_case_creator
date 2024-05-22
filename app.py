import pandas as pd
import json
import re
import unicodedata
from urllib.parse import quote

def sanitize_sheet_name(name):
    invalid_chars = '[]:*?/\\'
    for char in invalid_chars:
        name = name.replace(char, '')
    return name

def load_data(file):
    with open(file) as file:
        data = json.load(file)
    return data['workspace']['dialog_nodes'], data['workspace']['intents']

def process_node(node):
    node_title = node.get('title', '')
    node_conditions = node.get('conditions', 'No conditions')
    context = node.get('context', {})
    next_step = node.get('next_step')
    output = []
    response_type = None

    if 'output' in node and 'generic' in node['output']:
        for generic in node['output']['generic']:
            if 'values' in generic:
                output.extend(value.get('text', '') for value in generic['values'])
            if 'response_type' in generic:
                response_type = generic['response_type']

    output = '\n'.join(output) if len(output) > 1 else output[0] if output else ''

    behavior = None
    jump_to_node = None
    if 'next_step' in node:
        behavior = node['next_step'].get('behavior')
        if behavior == 'jump_to':
            jump_to_node = node['next_step'].get('dialog_node')

    intents = re.findall(r"#(\w+)", node_conditions)
    return (
        node_title, intents, str(context), next_step, behavior, jump_to_node, '\n'.join(output), response_type
    )
def process_intent(intent, dialog_nodes):
   
    nodes_by_intent_text = {}  
    intent_name = intent.get('intent', '')
    dialog_nodes.sort(key=lambda x: x.get('title', '') == 'No')

    if intent_name == "Bot_Control_Approve_Response":
        return nodes_by_intent_text  
    
    if intent_name == "Bot_Control_Reject_Response":
        return nodes_by_intent_text
    def follow_jump_to(node, visited_nodes=None):
       
        if visited_nodes is None:
            visited_nodes = set()

        dialog_node_id = node.get('dialog_node', '')
        if dialog_node_id in visited_nodes or node.get('title', '') == "Anything Else":
            return

        visited_nodes.add(dialog_node_id)

        title = node.get('title', '')
        output = node.get('output', {})
        generic = output.get('generic', [])
        context = node.get('context', {})
        if generic:
            for gen in generic:
                response_type = gen.get('response_type', '')
              
               
                if gen.get('values', []):
                    for value in gen['values']:
                        output_text = value.get('text', '')
                        nodes_by_intent_text[text].append([title, output_text , response_type])
        else:
            output_text = ''
            if isinstance(context, dict):
                if 'send_sms' in context and context['send_sms']:
                    output_text += 'Send SMS: ' + str(context['send_sms'])
                if 'sms_content' in context:
                    output_text += ' SMS Content: ' + context['sms_content']
            nodes_by_intent_text[text].append([title, output_text])

        next_step = node.get('next_step', {}).get('dialog_node', '')
        if next_step:
            next_node = next((node for node in dialog_nodes if node.get('dialog_node') == next_step), None)
            if next_node:
                follow_jump_to(next_node, visited_nodes)
        next_step_node = next((node for node in dialog_nodes if node.get('parent') == dialog_node_id), None)
        if next_step_node:
            follow_jump_to(next_step_node, visited_nodes)

        if node.get('behavior', '') == 'jump_to':
            jump_to_node = node.get('next_step', {}).get('dialog_node', '')
            jump_node = next((node for node in dialog_nodes if node.get('parent') == jump_to_node), None)
            if jump_node:
                follow_jump_to(jump_node, visited_nodes)
            
    intent_text = intent.get('text', '')
    examples = intent.get('examples', [])
    for example in examples:
        text = example.get('text', '')
        nodes_by_intent_text[text] = []  
        for i, node in enumerate(dialog_nodes):
            node_title, intents, context, next_step, behavior, jump_to_node, output, response_type = process_node(node)
            if intent_name in intents:
                response_type = output if intent_name == node_title else None
                output_text = output + str(context) if isinstance(output, str) else '\n'.join(output) + str(context)
                nodes_by_intent_text[text].append([intent_text, output_text, response_type])
                if behavior == 'jump_to':
                    for jump_node in dialog_nodes:
                        if 'parent' in jump_node and jump_node['parent'] == jump_to_node:
                            follow_jump_to(jump_node)
                else:
                    follow_jump_to(node)
    return nodes_by_intent_text

def write_to_excel(nodes_by_intent_text):
    with pd.ExcelWriter('dialog_skill.xlsx') as writer:
        for index, (text, rows) in enumerate(nodes_by_intent_text.items()):
            sanitized_sheet_name = sanitize_sheet_name(str(text)[:20]) 
            sanitized_sheet_name = f"{sanitized_sheet_name}_{index}"  
            df = pd.DataFrame(rows, columns=['Action','Expected Result AND Instructions for Next Steps', 'Response Type'])
            df = df.drop(columns=['Response Type'])  
            df.insert(0, 'Step #', '') 
            df.loc[-1] = ['', text, '']  
            df.index = df.index + 1 
            df = df.sort_index()  
            sanitized_sheet_name = sanitized_sheet_name.strip("'")
            df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)


def clean_entry(entry):
    if entry is None:
        return None
    cleaned_entry = entry.replace('<', '').replace('>', '').replace('Context: {}', '').replace('{', '').replace('}', '').replace('prosody rate="-25%"', '').replace("'initial_message': False", '').replace('/prosody', '').replace('break time="500ms"/', '').replace('express-as style="cheerful"', '').replace('prosody','').replace('break time="300ms"', '')
    cleaned_entry = unicodedata.normalize('NFKD', cleaned_entry).encode('ascii', 'ignore').decode('utf-8')
    cleaned_entry = cleaned_entry.replace('http', quote(cleaned_entry))
    cleaned_entry = re.sub(r'(\w)([A-Z])', r'\1 \2', cleaned_entry)
    return cleaned_entry
def dialog_skill(file):
    dialog_nodes, intents = load_data(file)
    nodes_by_intent_text = {}
    for intent in intents:
        nodes_by_intent_text.update(process_intent(intent, dialog_nodes))
    cleaned_nodes_by_intent_text = {k: [[clean_entry(item) for item in sublist] for sublist in v] for k, v in nodes_by_intent_text.items()}
    write_to_excel(cleaned_nodes_by_intent_text)

dialog_skill('voice-willow-dialog-v123.json')
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

    if 'output' in node and 'generic' in node['output']:
        for generic in node['output']['generic']:
            if 'values' in generic:
                output.extend(value.get('text', '') for value in generic['values'])

    output = node.get('output', {}).get('generic', [])
    output_text = []
    for item in output:
        for value in item.get('values', []):
            text = value.get('text', '')
            text = text.replace('\n', ' ').replace('<break time=\"500ms\"/>', ' ')  
            text = text.replace("W IC", "WIC")
            output_text.append(text)

    output_text = ' '.join(output_text)  # Join the output_text values into a single string
    node['output_processed'] = True

    behavior = None
    jump_to_node = None
    if 'next_step' in node:
        behavior = node['next_step'].get('behavior')
        if behavior == 'jump_to':
            jump_to_node = node['next_step'].get('dialog_node')

    intents = re.findall(r"#(\w+)", node_conditions)
    return (
        node_title, intents, str(context), next_step, behavior, jump_to_node, output_text
    )
def sort_dialog_nodes(dialog_nodes):
    node_indices = {id(node): index for index, node in enumerate(dialog_nodes)}
    dialog_nodes.sort(key=lambda x: (node_indices[id(x)], x.get('parent', ''), x.get('title', '') == 'No'))
    return dialog_nodes

def process_generic(node, nodes_by_intent_text, text):
    title = node.get('title', '')
    output = node.get('output', {})
    generic = output.get('generic', [])
    context = node.get('context', {})
    if generic and not node.get('output_processed', False):
        for gen in generic:
            if gen.get('values', []):
                output_texts = [value.get('text', '') for value in gen['values']]
                grouped_output_text = ' '.join(output_texts)
                if not any(grouped_output_text in sublist for sublist in nodes_by_intent_text[text]):
                    nodes_by_intent_text[text].append([title, grouped_output_text])
    # else:
    #     output_text = ''
    #     if isinstance(context, dict):
    #         if 'send_sms' in context and context['send_sms']:
    #             output_text = context['sms_content']
    #     if not any(output_text in sublist for sublist in nodes_by_intent_text[text]):
    #         nodes_by_intent_text[text].append([title, output_text])
    return nodes_by_intent_text
def follow_jump_to(node, dialog_nodes, nodes_by_intent_text, text, intent_text):
    def process_and_follow_jump(current_node, jump_source_title, visited_nodes):
        nonlocal nodes_by_intent_text
        current_dialog_node = current_node.get('dialog_node', '')
        if current_dialog_node in visited_nodes:
            return

        visited_nodes.add(current_dialog_node)

        node_title, intents, context, next_step, behavior, jump_to_node, output = process_node(current_node)
        if isinstance(context, dict):
            if 'send_sms' in context and context['send_sms']:
                output_text = context['sms_content']
            
        output_text = output + str(context) if isinstance(output, str) else '\n'.join(output) + str(context)
        if not any(output_text in sublist for sublist in nodes_by_intent_text[text]):
            nodes_by_intent_text[text].append([intent_text, output_text])
        
        jump_destination_title = current_node.get('title', '')

        child_nodes = [child for child in dialog_nodes if child.get('parent') == current_node.get('dialog_node', '')]
        yes_path = [child for child in child_nodes if child.get('title', '').lower() == "yes"]
        no_path = [child for child in child_nodes if child.get('title', '').lower() == "no"]
        anything_else = [child for child in child_nodes if child.get('title', '').lower() not in ["yes", "no"]]

        for child in yes_path:
            nodes_by_intent_text[text].append(["Choice: Yes", ""])
            
            if child.get('dialog_node', '') not in visited_nodes and child.get('title', '') not in ["Anything Else", "How else can I help you today?"]:
                process_and_follow_jump(child, jump_destination_title, visited_nodes.copy())

        for child in no_path:
            nodes_by_intent_text[text].append(["Choice: No", ""])
            if child.get('dialog_node', '') not in visited_nodes:
                jump_to_id = child.get('next_step', {}).get('dialog_node', '')
                if jump_to_id:  # If there's a jump_to ID
                    jump_node = next((node for node in dialog_nodes if node.get('dialog_node', '') == jump_to_id), None)
                    if jump_node:  # If a node with a matching dialog_node ID is found
                        process_and_follow_jump(jump_node, jump_destination_title, visited_nodes.copy())
                else:
                    process_and_follow_jump(child, jump_destination_title, visited_nodes.copy())
            else:
                nodes_by_intent_text[text].append(["Action: No choice selected", "Expected Result: No action taken"])

        for child in anything_else:
            if child.get('dialog_node', '') not in visited_nodes and child.get('title', '') not in ["Anything Else", "How else can I help you today?"]:
                nodes_by_intent_text[text].append([node_title])
                process_and_follow_jump(child, jump_destination_title, visited_nodes.copy())

    if node.get('title','') not in ["Anything Else", "How else can I help you today?"]:
        process_and_follow_jump(node, "", set())
    else:
        return nodes_by_intent_text
def process_intent(intent, dialog_nodes):
    nodes_by_intent_text = {}
    intent_name = intent.get('intent', '')
    dialog_nodes = sort_dialog_nodes(dialog_nodes)

    if intent_name in ["Bot_Control_Approve_Response", "Bot_Control_Reject_Response"]:
        return nodes_by_intent_text

    intent_text = intent.get('text', '')
    examples = intent.get('examples', [])
    for example in examples:
        text = example.get('text', '')
        nodes_by_intent_text[text] = []  
        visited_nodes = set()  # move the definition of visited_nodes here
        for i, node in enumerate(dialog_nodes):
            node_title, intents, context, next_step, behavior, jump_to_node, output = process_node(node)
            if intent_name in intents:
                output_text = output + str(context) if isinstance(output, str) else '\n'.join(output) + str(context)
                if intent_name != node_title:  
                    nodes_by_intent_text[text].append([intent_text, output_text])
                if behavior == 'jump_to':
                    for jump_node in dialog_nodes:
                        if 'parent' in jump_node and jump_node['parent'] == jump_to_node:
                            follow_jump_to(jump_node, dialog_nodes, nodes_by_intent_text, text, intent_text)
                else:
                    follow_jump_to(node, dialog_nodes, nodes_by_intent_text, text, intent_text)
    return nodes_by_intent_text
def write_to_excel(nodes_by_intent_text):
    with pd.ExcelWriter('dialog_skill.xlsx') as writer:
        for index, (text, rows) in enumerate(nodes_by_intent_text.items()):
            sanitized_sheet_name = sanitize_sheet_name(str(text)[:20]) 
            sanitized_sheet_name = f"{sanitized_sheet_name}_{index}"

            df = pd.DataFrame(rows, columns=['Action','Expected Result AND Instructions for Next Steps'])
            df.insert(0, 'Step #', '')
  
            df.loc[-1] = ['', text, '']  
            df.index = df.index + 1  
            df = df.sort_index()  
            sanitized_sheet_name = sanitized_sheet_name.strip("'")
            
            df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)

def clean_entry(entry):
    if entry is None:
        return None
    if isinstance(entry, bool):
        entry = str(entry)
    entry = entry.replace('\\xa0', ' ')
    cleaned_entry = entry.replace("<strong>", ' ').replace("</strong>", ' ').replace('<', '').replace('>', '').replace('Context: {}', '').replace('{', '').replace('}', '').replace('prosody rate="-25%"', '').replace("'initial_message': False", '').replace('/prosody', '').replace('break time="500ms"/', '').replace('express-as style="cheerful"', '').replace('prosody','').replace('break time="300ms"', '').replace('Â', ' ').replace('rate="-5%"', ' ').replace('/say-as', '').replace(' rate="-10%"', '').replace('break time="100ms"', '').replace('rate="-20%"', '').replace(' rate="-15%"', '').replace("'send_sms': True, 'sms_content': ", '').replace(' rate = "-25%"', '').replace("\n", '').replace("'ci_journey_step': 'Anything Else Help'", '')
    cleaned_entry = ''.join(c for c in cleaned_entry if c.isprintable())
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
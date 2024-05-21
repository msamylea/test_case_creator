import pandas as pd
import json
import re

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
    parent_node = node.get('parent')
    previous_sibling = node.get('previous_sibling')
    context = node.get('context', {})
    next_step = node.get('next_step')
    output = None
    if 'output' in node and 'generic' in node['output'] and len(node['output']['generic']) > 0 and 'values' in node['output']['generic'][0] and len(node['output']['generic'][0]['values']) > 0:
        output = node['output']['generic'][0]['values'][0]['text']

    if not output and 'sms_content' in context:
        output = context['sms_content']
    elif not output and 'text' in context:
        output = context['text']

    actions = node.get('actions')
    behavior = None
    jump_to_node = None
    if 'next_step' in node:
        behavior = node['next_step'].get('behavior')
        if behavior == 'jump_to':
            jump_to_node = node['next_step'].get('dialog_node')

    intents = re.findall(r"#(\w+)", node_conditions)
    return node_title, node_conditions, intents, parent_node, previous_sibling, context, next_step, actions, behavior, jump_to_node, output
def process_intent(intent, dialog_nodes):
    intent_name = intent.get('intent', '')
    examples = intent.get('examples', [])
    nodes_by_intent_text = {}
    for example in examples:
        text = example.get('text', '')
        if text not in nodes_by_intent_text:
            nodes_by_intent_text[text] = []
        for i, node in enumerate(dialog_nodes):
            node_title, node_conditions, intents, parent_node, previous_sibling, context, next_step, actions, behavior, jump_to_node, output = process_node(node)
            if intent_name in intents:
                response = output if intent_name == node_title else None
                if behavior == 'jump_to':
                    for jump_node in dialog_nodes:
                        if 'parent' in jump_node and jump_node['parent'] == jump_to_node:
                            context = jump_node.get('context', {})
                            if 'sms_content' in context:
                                response = context['sms_content']
                            elif 'text' in context:
                                response = context['text']
                nodes_by_intent_text[text].append([node_title, node_conditions, intents, parent_node, previous_sibling, context, next_step, actions, behavior, jump_to_node, response])
    return nodes_by_intent_text

def write_to_excel(nodes_by_intent_text):
    with pd.ExcelWriter('dialog_skill.xlsx') as writer:
        for index, (text, rows) in enumerate(nodes_by_intent_text.items()):
            sanitized_sheet_name = sanitize_sheet_name(str(text)[:20]) 
            sanitized_sheet_name = f"{sanitized_sheet_name}_{index}"  
            df = pd.DataFrame(rows, columns=['Title', 'Conditions', 'Intents', 'Parent', 'Previous Sibling', 'Context', 'Next Step', 'Actions', 'Behavior', 'Jump To Node', 'Response'])
            df.loc[-1] = [text, '', '', '', '', '', '', '', '', '', '']  
            df.index = df.index + 1 
            df = df.sort_index()  
            sanitized_sheet_name = sanitized_sheet_name.strip("'")
            df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)

def dialog_skill(file):
    dialog_nodes, intents = load_data(file)
    nodes_by_intent_text = {}
    for intent in intents:
        nodes_by_intent_text.update(process_intent(intent, dialog_nodes))
    write_to_excel(nodes_by_intent_text)

dialog_skill('voice-willow-dialog-v123.json')

def get_intents_texts(data):
    intents_texts = {}
    intents = data['workspace']['intents']

    for intent in intents:
        intent_name = intent['intent']
        examples = intent['examples']

        for example in examples:
            text = example['text']
            if intent_name in intents_texts:
                intents_texts[intent_name].append(text)
            else:
                intents_texts[intent_name] = [text]

    df = pd.DataFrame(list(intents_texts.items()), columns=['Intent', 'Texts'])
    df.to_excel('intents_text.xlsx', index=False)


def update_attachment(list_of_records):
    new_body= f"""[   {{"type": "TextBlock","text": "GVE SP Tracker (beta)","size": "Medium","color": "Dark"}},
    {{"type": "TextBlock","text": "Customer Name: {list_of_records[0]}", "size": "Medium","color": "Dark"}},
    {{"type": "TextBlock","text": "Project: {list_of_records[1]}", "size": "Medium","color": "Dark"}},
    {{"type": "TextBlock","text": "Category: {list_of_records[2]}", "size": "Medium","color": "Dark"}},
    {{"type": "Input.ChoiceSet",
                    "choices": [
                        {{
                            "title": "High",
                            "value": "High"
                        }},
                        {{
                            "title": "Medium",
                            "value": "Medium"
                        }},
                        {{
                            "title": "Low",
                            "value": "Low"
                        }}
                    ],
                    "placeholder": "Estimated complexity",
                    "id": "complexity",
                    "value": "{list_of_records[3]}"
                }},
                {{
                    "type": "Input.ChoiceSet",
                    "choices": [
                        {{
                            "title": "WiP",
                            "value": "WiP"
                        }},
                        {{
                            "title": "CE Pending",
                            "value": "CE Pending"
                        }},
                        {{
                            "title": "Close",
                            "value": "Close"
                        }}
                    ],
                    "placeholder": "Status",
                    "id": "status",
                    "value": "{list_of_records[4]}"
                }},     
            ]"""
    
    attachment =  {
    
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
            "type": "AdaptiveCard",
            "body": eval(new_body),
            "actions": [{
                    "type": "Action.Submit",
                    "title": "Update",
                    "data": "update",
                    "style": "positive",
                    "id": "button1"
                },
                {
                    "type": "Action.Submit",
                    "title": "Next",
                    "data": "next",
                    "style": "default",
                    "id": "button2"
                },
                {
                    "type": "Action.Submit",
                    "title": "Delete",
                    "data": "remove",
                    "style": "destructive",
                    "id": "button3"
                },
                {
                    "type": "Action.Submit",
                    "title": "Stop",
                    "data": "stop",
                    "style": "default",
                    "id": "button4"
                }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0"
            }
            }

    return attachment
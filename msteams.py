from pymsteams import connectorcard


class MsTeams:
    def __init__(self, webhook):
        self.teams = connectorcard(webhook)

    def create_mention_payload(self, text, user_id=None, user_display_name=None):
        adaptive_card_content = {
            "type": "AdaptiveCard",
            "body": [{"type": "TextBlock", "text": text, "wrap": True}],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0",
        }

        # メンションが指定されている場合、エンティティを追加
        if user_id and user_display_name:
            mention_entity = {
                "type": "mention",
                "text": f"<at>{user_display_name}</at>",
                "mentioned": {"id": user_id, "name": user_display_name},
            }
            adaptive_card_content["body"].insert(
                0,
                {
                    "type": "TextBlock",
                    "text": f"Hello {mention_entity['text']}",
                    "wrap": True,
                },
            )
            adaptive_card_content["msteams"] = {"entities": [mention_entity]}

        self.teams.payload = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": adaptive_card_content,
                }
            ],
        }

    def send(self):
        self.teams.send()

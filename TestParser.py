from html.parser import HTMLParser


class TestParser(HTMLParser):
    def __init__(self):
        self.tag_stack = []
        self.lst = []
        super().__init__()

    def error(self, message):
        print("ERROR: " + message)

    def handle_starttag(self, tag, attrs):
        self.tag_stack.insert(0, tag)

    def handle_endtag(self, tag):
        self.tag_stack.pop(0)

    def handle_data(self, data):
        data = data.strip()
        if data == "":
            return;
        if self.get_current_tag() == "li":
            self.lst += ["* " + data]
        else:
            self.lst += [data]

    def get_list(self):
        return self.lst

    def clear_list(self):
        self.lst = []

    def get_current_tag(self):
        if len(self.tag_stack) > 0:
            return self.tag_stack[0]
        else:
            return ""

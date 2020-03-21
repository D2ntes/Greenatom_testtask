class Context:
    pass


class Action:
    def process(self, context):
        return context

    def set_next(self, next_handler):
        self.next_handler = next_handler
        return next_handler

    def handle(self, context):
        self.process(context)
        if hasattr(self, 'next_handler'):
            context = self.next_handler.handle(context)
        return context

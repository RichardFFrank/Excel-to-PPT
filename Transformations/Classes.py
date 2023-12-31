class Slide:
    def __init__(self, slide_id: str, slide_type: str):
        self.slide_id = slide_id
        self.slide_type = slide_type
        self.content_locations = {}

    def add(self, content_location: str, content: str) -> bool:
        if content_location in self.content_locations.keys():
            self.content_locations[content_location] += [content]
        else:
            self.content_locations[content_location] = [content]
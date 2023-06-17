class Slide:
    
    def __init__(self, slide_id: str):
        self.slide_id = slide_id
        self.content_locations = {}


    def add(self, content_location: str, content: str) -> bool:
        if content_location in self.content_locations.keys():
            self.content_locations[content_location] += [content]
        else:
            self.content_locations[content_location] = [content]

if __name__ == "__main__":
    slide = Slide(3)
    slide.add( "%Title", "This is content")
    print(slide.content_locations)
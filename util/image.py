# Import required libraries
from PIL import Image

class ImageTooLargeError(Exception):
    def __init__(self, info: str) -> None:
        super().__init__(info)

class ImageData:
    "Used to manage and access the data of images."

    def __init__(self, directory: str, division_amount: int | float) -> None:
        # Load image with PIL
        self.image = Image.open(directory)
        # Image will be stretched - better to have a square image
        width = self.image.width / division_amount
        height = self.image.height / division_amount
        if height > 1048576 or width > 16384:
            raise ImageTooLargeError("The image provided is too large to be put into Excel.")
        self.image = self.image.resize((round(width), round(height)), Image.LANCZOS)

    def rgba_map(self) -> list[list[tuple[int, int, int, int]]]:
        """Returns a map with the RGBA values of every pixel within the image.
        Each new list is another layer of the image."""

        pix = self.image.load()
        fin = []
        p = -1
        for y in range(self.image.height-1):
            if y > p:
                fin.append([])
                p = y
            for x in range(self.image.width-1):
                fin[y].append(pix[x, y])
        return fin
        
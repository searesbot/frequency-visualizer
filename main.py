# Built on a Windows machine, for Windows machines

# For creating GUI which user interacts with to run this program
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
import os

import docx # For interacting with Microsoft Word documents

# For removing stopwords & creating visualizations from text
from nltk.corpus import stopwords
from wordcloud import WordCloud
import matplotlib.pyplot as plotter
from PIL import ImageTk, Image

win = Tk() # Create an instance of Ttkinter frame [this is the GUI]

stopwords = stopwords.words("english") # Declare stopwords using NLTK's pre-determined list
uniform_width = 550 # Set width of all graphic elements in program
uniform_height = 325 # Set height of all graphic elements in program

def get_text(path: str):
    # TODO: Add support for other file extensions
    # Save text in .docx file at path to string variable (to pass into instance of word cloud) & return the string
    file = docx.Document(path)
    text: str = ""
    for paragraph in file.paragraphs:
        text += paragraph.text.lower() # Convert text in each paragraph to lowercase and add to persistent string
    return text

def create_wordcloud(text: str):
    # Create a word cloud graphic (will be converted into an image)
    wordcloud = WordCloud(width = uniform_width, height = uniform_height, max_words = 25, stopwords = stopwords, mode = "RGBA", relative_scaling = 1.0, background_color = None, collocations = False, regexp = None, min_word_length = 4, collocation_threshold = 100)
    wordcloud.generate(text) # Pass file into word cloud instance
    return wordcloud

def create_image(wc: WordCloud):
    raw_image = wc.to_image() # Create image using WordCloud's in-built method
    processed_image = ImageTk.PhotoImage(raw_image) # Convert image into form usable in GUI
    return processed_image

def update_gui(visualization: PhotoImage):
    # Show visualization on-screen
    canvas = Canvas(win, width = uniform_width, height = uniform_height)
    canvas.pack() # Add a Canvas to the GUI
    image_container = canvas.create_image(0, 0, anchor = "nw", image = visualization)
    # TODO: Debug PhotoImage TypeError caused by next line
    canvas.itemconfig(image_container, visualization) # Add the word cloud image into the Canvas
    # Manual call to GUI's update methods to show image on-screen
    win.update_idletasks()
    win.update()

def process(document_path):
    text = get_text(document_path) # Convert contents of target file to a string
    wc = create_wordcloud(text) # Create a word cloud instance using the returned string
    image = create_image(wc) # Convert the word cloud instance into an image
    update_gui(image) # Show the word cloud on the GUi

def open_file():
    # Prompt user to browse local file system for target file
    file = filedialog.askopenfile(mode = "r")
    if file:
        # If a file has been chosen, internally store the absolute path of the file
        document_path = os.path.abspath(file.name)
        top_label.pack_forget() # Remove label widget
        browse_button.pack_forget() # Remove button widget
        process(document_path) # Call a control method that will govern the rest of the program's execution

if __name__ == "__main__":

    win.geometry("{}x{}".format(uniform_width, uniform_height))
    win.title("document term strength visualizer")

    # Add Label widgets & Button
    top_label = Label(win, text="Click 'Browse' to select file from local machine")
    browse_button = ttk.Button(win, text="Browse", command = open_file)
    # TODO: Add arrow button so user can choose number of words to show in visualization
    top_label.pack()
    browse_button.pack()

    win.mainloop() # Initialize GUI & begin presenting information to user
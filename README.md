# Comparing two PowerPoint pptx files to find similar slides

This Python tool compares two PowerPoint pptx files, searching for similar slides based on two criteria: image similarity and text similarity.

It generates a report like the below report:

![kioku-250516-191130-2633](https://github.com/user-attachments/assets/47ba0bd6-cc84-4b09-bda7-95dfb6cdc58c)

This tool can help when you have many subtly different pptx files as a result of repeated fine-tuning.

## Requirements

### Operating System

This tool requires PowerPoint to function and works on Windows 10 or later.
(It has not been tested on macOS.)

### Required Software

Before running the project, make sure you have the following software installed:

- Python 3.7 or higher
- Microsoft PowerPoint for Windows

### Additional Tools

- Web browser (needed to view the result report).

## How to use the tool

### 1. install required packages

To install the required Python packages, run the following command:

```bash
pip install ImageHash numpy scikit-learn comtypes Pillow sentence-transformers
```

### 2. run the Script

After modifying Oldslides.pptx to create Newslides.pptx, you may want to see what has changed. In that case, you can find the modifications using the following:

```bash
python compare-pptx.py Newslides.pptx Oldslides.pptx
```

"compare-pptx.py" generates a report that lists slides in Oldslides.pptx that are similar to those in Newslides.pptx.

### 3. check the report

The report will be generated in the export/analyzed#DATETIME#/comparison\_report.html in the current directory. #DATETIME# will be replaced with the timestamp of when the script is run.

## the format of the report

- Original: lists all slides of Newslides.pptx, including no matches found.
- Match: lists slides of Oldslides.pptx that matched with the original (meaning almost identical). Multiple slides can be listed.
- High: lists slides of Oldslides.pptx, with  high similarity. Multiple slides can be listed.
- Low: lists slides of Oldslides.pptx, with  low similarity. This means that there is a slight possibility that it could be a similar slide. Multiple slides can be listed.
- ImageDifference: this dark picture shows the difference between the old and the new. Parts of the two images that have no difference will appear black, while the areas with differences will appear in brighter colors. This makes it easy to identify the areas with differences.
- NewSlide: the file name of the image output of the corresponding slide from Newslides.pptx
- ImageScore: the similarity score of the images of the two slides. A small value means high similarity, zero means almost identical.
- TextScore: the similarity score of the texts of the two slides. A small value means high similarity, zero means almost identical.

![kioku-250516-172010-2632](https://github.com/user-attachments/assets/a56ca9bc-4655-4965-b5ee-cd90d596f6b3)

## How it works

The overall process flow looks like this:

![kioku-250517-162036-2639](https://github.com/user-attachments/assets/744c305c-df54-4e10-93e3-efd8cb65ca54)

### exporting images of the pptx files

Assuming that you have two pptx, Newslides.pptx and Oldslides.pptx, specify them as the parameters to the compare-pptx.py script.
The script exports images of them to the deriveddir and the basedir. The new is a derivative of the old, so the working directories are called as such.

### calculate hash values

The script calculates the imagehash and the textvector of the images, and stores them in derived.json and base.json. An imagehash is a hash value that summarizes the visual characteristics of an image, useful for estimating the similarity of images. The value is calculated using the following formula:

```python
import imagehash
with Image.open(imagepath) as img:
    hash = imagehash.phash(img)
```

A textvector is also a summarized value of the characteristics of the text of a slide. A slide contains multiple shapes with text, so the script extracts and concatenates them to a long text, then calculates the textvector using the following formula.

```python
model = SentenceTransformer('all-MiniLM-L6-v2') 
textvector = model.encode(slide_text)
```

Calculated hash values are stored in json files in the exportroot directory, such as: derived_analyzed.json and base_analyzed.json.

The script invokes the PowerPoint application to export images and collect text from slides.

### compares hash values, finds similarities and stores them

The script compares the hash values of the new and the old slides one by one, finds similarity and stores the information of the pair of high similarity in the derived_analyzed.json file. A slide of the new pptx could have multiple slides of the new as the similars.

### reports the similarities as an HTML file

Finally, the script generates a report of the similarities as an html file, named comparison_report.html in the exportroot directory.

## clean up

The script makes working directories and files under the exportroot directory, and you can specify the location by the --exportroot option, such as:

```bash
python compare-pptx.py --exportroot DIR Newslide.pptx Oldslide.pptx
```

where DIR can be a relative/absolute path. If it contains '#DT#' string, it will be replaced with the current datetime string. The default exportroot is './export/analyzed#DT#'.

After you are done, you can delete the entire working directories generated.

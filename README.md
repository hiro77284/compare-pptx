# Comparing two PowerPoint pptx files to find similar slides.

This Python tool compares two PowerPoint pptx files, searching similar slides based on two criteria: image similarity and text similarity.

It generates the report looks like below:

![kioku-250516-191130-2633](https://github.com/user-attachments/assets/47ba0bd6-cc84-4b09-bda7-95dfb6cdc58c)

It helps you when you have many subtly different pptx files as a result of repeated fine-tuning of a pptx.

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

## How to use

### install required packages

To install the required Python packages, run the following command:

```bash
pip install ImageHash numpy scikit-learn comtypes Pillow sentence-transformers
```

### running the Script

After modifying Oldslides.pptx to create Newslides.pptx, you may want to find the changes. In that case, you can find the modifications using the following command.

```bash
python compare-pptx.py Newslides.pptx Oldslides.pptx
```

"compare-pptx.py" generates a report that lists similar slides from Oldslides.pptx for each slide in Newslides.pptx.

### the location of the report

The report will be generated in the export/analyzed#DATETIME#/comparison\_report.html under the current directory. #DATETIME# will be replaced with the timestamp of when the script is run.

## the format of the report

- Original: lists all slides of Newslides.pptx, including no matches found.
- Match: lists slides of Oldslides.pptx, matched with the original, it means almost identical. Multiple slides can be listed, if any.
- High: lists slides of Oldslides.pptx, having high similarity with the original. Multiple slides can be listed, if any.
- Low: lists slides of Oldslides.pptx, having relatively low similarity with the original. It means that there is a slight possibility that it could be a similar slide. Multiple slides can be listed, if any.
- ImageDifference: this almost black picture shows the diffrerence between the old and the new. Parts of the two images that have no difference will appear black, while the areas with differences will appear in brighter colors. This makes it easy to identify the areas with differences.
- NewSlide: the file name of the image output of the corresponding slide from Newslides.pptx
- ImageScore: the similarity score of the images of the two slides. Small value means high similarity. Zero means almost identical.
- TextScore: the similarity score of the texts of the two slides. Small value means high similarity, Zero means almost identical.

![kioku-250516-172010-2632](https://github.com/user-attachments/assets/a56ca9bc-4655-4965-b5ee-cd90d596f6b3)

## How it works

The overall process flow looks like this:

### exporting images of the pptx files

Assuming that you have two pptx that Newslides.pptx and Oldslides.pptx, specify them as the parameters to the compare-pptx.py script.
The script exports images of them to the driveddir and the basedir. The new is a derivative of the old, so the working directories are called as such.

### calculate hash values

The script calculates the imagehash and the textvector of the images, stores in derived.json and base.json. An imagehash is a hash value that summarizes the visual characteristics of an image, useful for estimating the similarity of images. The value is calculated using the following formula.

```python
import imagehash
with Image.open(imaagepath) as img:
    hash = imagehash.phash(img)
```

A textvector is also a summarized value of the characteristics of text of a slide. A slide contains multiple shapes with texts, so the script extracts and concatinates them to a long text, calculates the textvector using the following formula.

```python
model = SentenceTransformer('all-MiniLM-L6-v2') 
textvector = model.encode(slide_text)
```

Calculated hash values are stored in json files in the exportroot directory, such as: derived_analyzed.json and base_analyzed.json.

The script invokes the PowerPoint application to export images and collect text from slides.

### compare hash values, find similarity, assigns new and old slides

The script compares the hash values of the new and the old slides one by one, finds similarity and stores the information of the pair of high similarity in the derived_analyzed.json file. A slide of the new pptx could have multiple slides of the new as the similars.

### repots the assignments as an html file

Finally, the script generates a report of the assignments as an html file, named comparison_report.html in the exportroot directory.

## clean up

The script makes working directories and files under the exportroot directory, you can specify the location by the --exportroot option, such as:

```bash
python compare-pptx.py --exportroot DIR Newslide.pptx Oldslide.pptx
```

where DIR can be a relative/absolute path. If it contains '#DT#' string, it would be replace to the current datetime string. The default exportroot is './export/analyzed#DT#'.
![kioku-250516-191130-2633](https://github.com/user-attachments/assets/47ba0bd6-cc84-4b09-bda7-95dfb6cdc58c)


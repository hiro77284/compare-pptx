# Comparing two PowerPoint pptx files to find similar slides.

This Python tool compares two PowerPoint pptx files, searching similar slides based on two criteria: image similarity and text similarity.

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
- ImageScore: the similarity score of the two images. Small value means high similarity. Zero means almost identical.
- TextScore: the similarity score of the texts of the two slides. Large value means high similarity, 100 means almost identical.

![kioku-250516-172010-2632](https://github.com/user-attachments/assets/a56ca9bc-4655-4965-b5ee-cd90d596f6b3)

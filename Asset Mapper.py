import numpy as np
import pandas
import folium
from folium import IFrame
import re
import os
import math

# TK will eventually handle this input...
configpath = "config.xlsx"


# overlaypath = "resources\\images\\Tunneloverlay.jpg"  # apparently it takes a filename not a file handle... neat!

class assetclass:
    def __init__(self, name, latitude, longitude, desc, spherepath, imagepath):
        self.name = name
        self.latitude = latitude
        self.longitude = longitude
        self.desc = desc
        self.spherepath = spherepath
        self.imagepath = imagepath

# Set up list and data class for locations
marker_list = []
# Parse some data from the excel file
print("Reading Config Excel File...")
configdata = pandas.read_excel(configpath)
for index, row in configdata.iterrows():
    excelfilepath = row['DataPath']
    popupheight = row['PopupHeight']
    popupwidth = row['PopupWidth']
    overlaypath = row['OverlayPath']
    OverlayN = row['OverlayN']
    OverlayS = row['OverlayS']
    OverlayE = row['OverlayE']
    OverlayW = row['OverlayW']
    sitebounds = [[OverlayN, OverlayW], [OverlayS, OverlayE]]  # these are the bounds of the overlay image.
    StartZoom = row['StartZoom']
print("Done!")

print("Reading Asset Data Store...")
exceldata = pandas.read_excel(excelfilepath)
print("Opened Excel File")

for index, row in exceldata.iterrows():
    if str(row['Name']) == "nan":  # skip nan rows as they throw lots of nice little errors
        continue
    newasset = assetclass(row['Name'], row['Latitude'],row['Longitude'], row['Desc'], row['Spherepath'], row['Imagepath'])

    # Store new asset in dictionary, using asset tag as lookup
    marker_list.append(newasset)
print("Created marker dictionary from Asset List.")
# Display Full Readout of Marker Data upon program finish
print("Asset list:")
for i in marker_list:
    print(
        f"{i.name} at {i.latitude},{i.longitude} with image = {i.spherepath}")

# Time to make the folium map.
assetmap = folium.Map(location=[(OverlayN + OverlayS) / 2, (OverlayW + OverlayE) / 2], zoom_start=StartZoom,
                      max_zoom=1000)  # centered map on site.
img = folium.raster_layers.ImageOverlay(
    name="Mascot Tunnel Test Map Overlay",
    image=overlaypath,
    bounds=sitebounds,
    opacity=1,
    interactive=True,
    cross_origin=False,
    zindex=1,
)
img.add_to(assetmap)

A = folium.LatLngPopup()  # Allows us to browse the map and get the location of new assets. Roughly at least...
A.add_to(assetmap)

for currentlocation in marker_list:  # for each location that we have need a marker on...
    # make the marker on the map
    html = f"""
    <iframe src=resources/sphere{currentlocation.name}.html height={popupheight} width={popupwidth}></iframe>
     """
    folium.Marker(
        show=True,
        location=[currentlocation.latitude, currentlocation.longitude],
        popup=folium.Popup(html=html, sticky=True),
        tooltip=currentlocation.name,
        sticky=True,
        closePopupOnClick=False,
        autoClose=False,
    ).add_to(assetmap)

    # make the corresponding photosphere file.

    newsphere = open(f'resources/sphere{currentlocation.name}.html', "w")
    newsphere.write(f"""<link rel="stylesheet" href="photo-sphere-viewer.min.css"/>
<script src="three.min.js"></script>
<script src="browser.min.js"></script>
<script src="photo-sphere-viewer.min.js"></script>
{currentlocation.name} - {currentlocation.desc}

<div id="viewer"></div>

<style>
  /* the viewer container must have a defined size */
  #viewer {{
    width: 95vw;
    height: 80vh;
    margin: auto;
  }}
</style>

<script>
  var viewer = new PhotoSphereViewer.Viewer({{
    container: document.querySelector('#viewer'),
    panorama: 'images/{currentlocation.spherepath}'
  }});
</script>
<center><a href=sphere{currentlocation.name}.html target="_blank"><centre>Open in New Tab</a></center>""")
    newsphere.close()

print("Saving map....", end='')
assetmap.save('Map.html')
print(" Map saved in working directory!")

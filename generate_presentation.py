import os
import pandas as pd
import re
from datetime import datetime

base_path = os.getcwd()
rushee_data_path = os.path.join(base_path, 'rush_responses.xlsx')
image_dir = os.path.join(base_path, 'rushee_images')
presentation_path = os.path.join(base_path, 'presentation.html')

def get_rushee_info_dict(data_path):
  contacts = []
  data_df = pd.read_excel(data_path).fillna('').sort_values(by=['Rushee Name'])
  contacts = list(dict(contact_row[1][1:]) for contact_row in data_df.iterrows())
  rushees = dict()
  for contact in contacts:
    rushee_name = contact['Rushee Name']
    rusher_name = contact['Your Name']
    rusher_excitement = (rusher_name, contact['How excited are you about them?'])
    comment = contact['Rushee Information']
    if rushee_name not in rushees:
      rushees[rushee_name] = {
          'rusher excitement': [],
          'comments': []
      }
    rushee_dict = rushees[rushee_name]
    rushee_dict['rusher excitement'].append(rusher_excitement)
    rushee_dict['primary'] = contact['Primary']
    rushee_dict['secondary'] = contact['Secondary']
    rushee_dict['bucket'] = contact['Bucket']
    rushee_dict['cross rush'] = contact['Cross-Rush']
    rushee_dict['closers'] = contact['Closers']
    rushee_dict['year'] = contact['Year']
    if comment:
      rushees[rushee_name]['comments'].append(comment)
  return rushees

def addSlide(primary, secondary, rushee, comments, photoURL, closers = "", class_year = 2025, bucket = None, cross_rush = None):
    """Returns an html markup of the rushee slide
    """    
    comments_string = ''
    for comment in comments:
      comments_string += "<li>" + comment + "</li>"

    if not bucket:
      bucket = 'N/A'

    if not cross_rush:
      cross_rush = '???'
        
    return ("""
              <section>
                <div id='slide'>
                    <div id='name'>
                        <h1>{rushee}</h1>
                    </div>
                    <div class='flex-container'>
                        <div id='profile'>
                            <div id='pic'>
                                <img src="{photoURL}" />
                            </div>
                            <div id='info'>
                                <table>
                                    <tr>
                                        <td>Primary/Secondary</td>
                                        <td>{primary}/{secondary}</td>
                                    </tr>
                                    <tr>
                                        <td>Closers</td>
                                        <td>{closers}</td>
                                    </tr>
                                    <tr>
                                      <td> Class Year </td>
                                      <td> {class_year} </td>
                                    </tr>
                                    <tr>
                                        <td>Bucket</td>
                                        <td>{bucket}</td>
                                    </tr>
                                    <tr>
                                        <td> x-rush</td>
                                        <td> {cross_rush} </td>
                                    </tr>
                                </table>
                            </div>
                        </div>
                        <div id='about'>
                            <div id='about1'>
                            <ul>
                                {comments}
                            </ul>
                            </div>
                            <div id='about2'>
                            <ul>
                            </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </section>
        """.format(photoURL = photoURL,
                    primary = primary,
                    secondary = secondary,
                    rushee = rushee,
                    bucket = bucket,
                    comments = comments_string,
                    cross_rush = cross_rush,
                    class_year = class_year,
                    closers = closers
                    )
    )

default_image_path = os.path.join(image_dir, 'default.jpg')

def get_image_path(rushee):
    rushee_photo_url = os.path.join(image_dir, '_'.join(rushee.rstrip().lower().split()))
    print(rushee, " ", rushee_photo_url)
    if os.path.exists(rushee_photo_url + '.jpg'):
      return rushee_photo_url + '.jpg'
    elif os.path.exists(rushee_photo_url + '.png'):
      return rushee_photo_url + '.png'
    else:
      return default_image_path
  
def write_to_slides(rushee_info_dict, presentation_path):
  print(f'Writing to: {presentation_path}')
  with open(presentation_path, 'w', encoding="utf-8") as f:
      f.write(r"""
          <html>
              <head>
                  <link rel="stylesheet" href="assets/css/reveal.css">
                  <link rel="stylesheet" href="assets/css/theme/black.css">
                  <link rel="stylesheet" href="assets/css/custom.css">
              </head>
              <body>
                  <div class="reveal">
                      <div class="slides">
      """)
      for rushee_name, rushee_info in rushee_info_dict.items(): 
          rushee_slide = addSlide(
              primary = rushee_info['primary'], 
              secondary = rushee_info['secondary'], 
              rushee = rushee_name,   
              comments = rushee_info['comments'], 
              class_year = rushee_info['year'], 
              bucket = rushee_info['bucket'], 
              photoURL = get_image_path(rushee_name), 
              closers = rushee_info['closers']
            )
          # write body contents to file
          f.write(rushee_slide)
      # write footer to html file
      f.write(r"""
                      </div>
                  </div>
                  <script src="assets/js/reveal.js"></script>
                  <script>
                      Reveal.initialize({
                          controls: true,
                          progress: true,
                          history: true,
                          center: true,
                          slideNumber: true,
                          transition: 'slide' // none/fade/slide/convex/concave/zoom
                          });
                  </script>
              <body>
          </html>
      """)

def generate_presentation(data_path = rushee_data_path, presentation_path = presentation_path):
  rushee_info_dict = get_rushee_info_dict(data_path = data_path)
  write_to_slides(rushee_info_dict, presentation_path = presentation_path)

if __name__ == '__main__':
    generate_presentation()

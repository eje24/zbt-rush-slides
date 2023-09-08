import os
import pandas as pd
from typing import List, Dict, NamedTuple

base_path = os.getcwd()
rushee_data_path = os.path.join(base_path, "rush_responses.xlsx")
image_dir = os.path.join(base_path, "rushee_images")
default_image_path = os.path.join(image_dir, "default.jpg")
presentation_path = os.path.join(base_path, "presentation_test.html")


class RusheeInfo(NamedTuple):
    comments: List[str]
    primary: str
    bucket: str
    closers: List[str]
    status: str


class RusheeInfoDict(Dict):
    def __init__(self):
        super().__init__()

    def update(self, rush_response):
        name = rush_response["Rushee Name"].lower().rstrip()
        comment = rush_response["Rushee Information"]
        primary = rush_response["Primary"]
        bucket = rush_response["Bucket"]
        closers = rush_response["Closers"]
        status = rush_response["Status"]

        if name not in self:
            self[name] = RusheeInfo(
                comments=[],
                primary="",
                bucket="",
                closers="",
                status="",
            )

        self[name].comments.append(comment)
        self[name] = RusheeInfo(
            comments=self[name].comments,
            primary=primary if self[name].primary == "" else self[name].primary,
            bucket=bucket if self[name].bucket == "" else self[name].bucket,
            closers=closers if self[name].closers == "" else self[name].closers,
            status=status if self[name].status == "" else self[name].status,
        )

    def generate(data_path):
        rushee_info_dict = RusheeInfoDict()
        data_df = pd.read_excel(data_path).fillna("")
        updates = list(dict(contact_row[1][1:]) for contact_row in data_df.iterrows())
        for update in updates:
            rushee_info_dict.update(update)
        return rushee_info_dict


def addSlide(name, comments, primary, bucket, closers, status, photoURL):
    """Returns an html markup of the rushee slide"""
    comments_string = ""
    for comment in comments:
        comments_string += "<li>" + comment + "</li>"

    if not bucket:
        bucket = "N/A"

    return """
              <section>
                <div id='slide'>
                    <div id='name'>
                        <h1>{name}</h1>
                    </div>
                    <div class='flex-container'>
                        <div id='profile'>
                            <div id='pic'>
                                <img src="{photoURL}" />
                            </div>
                            <div id='info'>
                                <table>
                                    <tr>
                                        <td>Primary</td>
                                        <td>{primary}</td>
                                    </tr>
                                    <tr>
                                        <td>Bucket</td>
                                        <td>{bucket}</td>
                                    </tr>
                                    <tr>
                                        <td>Closers</td>
                                        <td>{closers}</td>
                                    </tr>
                                    <tr>
                                        <td>Status</td>
                                        <td>{status}</td>
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
        """.format(
        name=name,
        photoURL=photoURL,
        primary=primary,
        bucket=bucket,
        comments=comments_string,
        status=status,
        closers=closers,
    )




def get_image_path(rushee):
    rushee_photo_url = os.path.join(
        image_dir, "_".join(rushee.rstrip().lower().split())
    )
    if os.path.exists(rushee_photo_url + ".jpg"):
        return rushee_photo_url + ".jpg"
    elif os.path.exists(rushee_photo_url + ".jpeg"):
        return rushee_photo_url + ".jpeg"
    elif os.path.exists(rushee_photo_url + ".png"):
        return rushee_photo_url + ".png"
    else:
        return default_image_path


def write_to_slides(rushee_info_dict, presentation_path):
    with open(presentation_path, "w", encoding="utf-8") as f:
        f.write(
            r"""
          <html>
              <head>
                  <link rel="stylesheet" href="assets/css/reveal.css">
                  <link rel="stylesheet" href="assets/css/theme/black.css">
                  <link rel="stylesheet" href="assets/css/custom.css">
              </head>
              <body>
                  <div class="reveal">
                      <div class="slides">
      """
        )
        for rushee_name, rushee_info in rushee_info_dict.items():
            if rushee_info.bucket == "Drop":
                continue
            rushee_slide = addSlide(
                name=rushee_name,
                comments=rushee_info.comments,
                primary=rushee_info.primary,
                bucket=rushee_info.bucket,
                closers=rushee_info.closers,
                status=rushee_info.status,
                photoURL=get_image_path(rushee_name),
            )
            # write body contents to file
            f.write(rushee_slide)
        # write footer to html file
        f.write(
            r"""
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
      """
        )


def generate_presentation(
    data_path=rushee_data_path, presentation_path=presentation_path
):
    rushee_info_dict = RusheeInfoDict.generate(data_path)
    write_to_slides(rushee_info_dict, presentation_path=presentation_path)


if __name__ == "__main__":
    generate_presentation()

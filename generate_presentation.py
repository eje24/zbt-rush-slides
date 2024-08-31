import os
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Tuple

import pandas as pd

base_path = os.getcwd()
rushee_data_path = os.path.join(base_path, "rush_responses.xlsx")
image_dir = os.path.join(base_path, "rushee_images")
default_image_path = os.path.join(image_dir, "default.jpg")
presentation_path = os.path.join(base_path, "presentation.html")


class Bucket(Enum):
    NONE = ""
    DROP = "Drop"
    PULL = "Pull"
    PASS = "Pass"

    @classmethod
    def max(cls, old_bucket: "Bucket", new_bucket: "Bucket") -> "Bucket":
        return new_bucket if new_bucket.name > old_bucket.name else old_bucket

    @classmethod
    def valid(cls, bucket_name: str) -> bool:
        return bucket_name in [b.value for b in cls]


@dataclass(frozen=True)
class RusheeInfo:
    """
    Information about a rushee
    """

    name: str
    comments: Tuple[str, ...]
    primary: str
    bucket: Bucket
    closers: Tuple[str, ...]
    status: str

    @classmethod
    def from_rush_response(cls, rush_response: Dict[str, Any]) -> "RusheeInfo":
        comment = rush_response["Rushee Information"].strip()
        return cls(
            name=rush_response["Rushee Name"].lower().rstrip(),
            comments=(comment,) if comment else tuple(),
            primary=rush_response["Primary"],
            bucket=(
                Bucket(rush_response["Bucket"])
                if Bucket.valid(rush_response["Bucket"])
                else Bucket.NONE
            ),
            closers=rush_response["Closers"],
            status=rush_response["Status"],
        )

    def merge(self, other: "RusheeInfo") -> "RusheeInfo":
        return RusheeInfo(
            name=self.name,
            comments=self.comments + other.comments,
            primary=self.primary or other.primary,
            bucket=Bucket.max(self.bucket, other.bucket),
            closers=self.closers or other.closers,
            status=self.status or other.status,
        )


class AggregateRusheeInfo:
    def __init__(self):
        self.info: Dict[str, RusheeInfo] = {}

    def update(self, response_info: RusheeInfo):
        name = response_info.name
        info = self.info
        if name not in info:
            info[name] = response_info
        else:
            info[name] = info[name].merge(response_info)

    @classmethod
    def parse_df(cls, rush_responses_df) -> "AggregateRusheeInfo":
        aggregate_rushee_info = cls()
        rush_responses = list(
            dict(contact_row[1][1:]) for contact_row in rush_responses_df.iterrows()
        )
        for rush_response in rush_responses:
            response_info = RusheeInfo.from_rush_response(rush_response)
            aggregate_rushee_info.update(response_info)
        return aggregate_rushee_info


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


def write_to_slides(info):
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
        for rushee_name, info in info.info.items():
            if info.bucket == "Drop":
                continue
            rushee_slide = addSlide(
                name=rushee_name,
                comments=info.comments,
                primary=info.primary,
                bucket=info.bucket.name,
                closers=info.closers,
                status=info.status,
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


def generate_presentation(data_path=rushee_data_path):
    rush_responses_df = pd.read_excel(data_path).fillna("")
    rushee_info = AggregateRusheeInfo.parse_df(rush_responses_df)
    write_to_slides(rushee_info)


if __name__ == "__main__":
    generate_presentation()

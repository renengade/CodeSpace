import os, sys
from PIL import Image


def pics_path(res_name):
    if not os.path.exists(f"{res_name}"):
        os.makedirs(res_name)
        print("Directories Generated !!!")
    else:
        print("Directories Already Exists !!!")


# 批量处理
def reduce_size(w_scale=0.5, h_scale=0.5):
    for pic in pics:
        try:
            im = Image.open(pic)
            w, h = im.size
            img = im.resize(
                (int(w * w_scale), int(h * h_scale)), Image.Resampling.LANCZOS
            )
            img.save(f"调整尺寸\\{pic}")
        except Exception as e:
            print(f"errors:{e} occour when reducing the size of pictures!!!")


if __name__ == "__main__":
    wpath = os.path.dirname(os.path.abspath(__file__))
    os.chdir(wpath)
    pics = [x for x in os.listdir() if x.endswith("jpg") or x.endswith("png")]
    pics_path("调整尺寸")
    reduce_size()

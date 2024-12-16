import pandas as pd
from py3dbp import Packer, Bin, Item, Painter
import time
start = time.time()
# data = pd.read_excel("C:/Users/vuong/OneDrive/Máy tính/Input.xlsx")

import tkinter as tk
from tkinter import filedialog

# Tạo cửa sổ Tkinter
root = tk.Tk()
root.withdraw()  # Ẩn cửa sổ chính của Tkinter

# Hiển thị hộp thoại chọn tệp
file_path = filedialog.askopenfilename(title="Chọn tệp Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])

# Kiểm tra xem người dùng có chọn tệp không
if file_path:
    try:
        # Đọc tệp Excel
        data = pd.read_excel(file_path)
        print("Đã tìm thấy file")
        print("******************************")
        print(f"Số hàng: {data.shape[0]}, Số cột: {data.shape[1]}")
        print("Dữ liệu đầu tiên trong tệp:")
        print(data.head())
    except Exception as e:
        print(f"Có lỗi xảy ra: {e}")
else:
    print("Không có tệp nào được chọn.")
'''

If you have multiple boxes, you can change distribute_items to achieve different packaging purposes.
1. distribute_items=True , put the items into the box in order, if the box is full, the remaining items will continue to be loaded into the next box until all the boxes are full  or all the items are packed.
2. distribute_items=False, compare the packaging of all boxes, that is to say, each box packs all items, not the remaining items.

'''

# init packing function
packer = Packer()
# Khởi tạo danh sách các bins
# bins = [
#     Bin(f'Bin{i+1}', (1100, 1100, 900), 10000, 0, 0)
#     for i in range(20)
# ]

# Nhập số lượng bin
num_bins = int(input("Nhập số lượng bin: "))

# Nhập kích thước bin
width = float(input("Nhập chiều dài: "))
height = float(input("Nhập chiều rộng: "))
depth = float(input("Nhập chiều cao: "))
print("Đã nhận dữ liệu. Đang tiến hành sắp xếp ...")

# Tạo danh sách các bin
bins = [
    Bin(f'Bin{i+1}', (width, height, depth), max_weight=10000)
    for i in range(num_bins)
]

# Thêm các bins vào packer
for bin_item in bins:
    packer.addBin(bin_item)


colors = [
    "Red", "Green", "Blue", "Yellow", "Orange", "Purple", "Pink", "Brown", "Gray", "Black",
    "White", "Violet", "Indigo", "Cyan", "Magenta", "Turquoise", "Gold", "Silver", "Beige", "Lavender",
    "Coral", "Salmon", "Teal", "Maroon", "Olive", "Lime", "Crimson", "Emerald", "Aqua", "Peach",
    "Tan", "Mint", "Plum", "Rose", "Azure", "Chartreuse", "Amber", "Mahogany", "Ivory", "Cerulean",
    "Slate", "SeaGreen", "SkyBlue", "ElectricBlue", "Periwinkle", "MintGreen", "Saffron", "Blush", "Fuchsia",
    "TurquoiseBlue", "Tangerine", "Sunset", "Onyx", "Jade", "Mulberry"
]

num_rows = data.shape[0]

# Tạo mảng có số lượng phần tử bằng số hàng trong DataFrame
arr = [0] * num_rows  # Tạo mảng có số lượng phần tử bằng số hàng, giá trị mặc định là None
# print (len(array))

# Duyệt qua từng dòng của DataFrame
for index, row in data.iterrows():
    quantity = int(row['Quantity'])  # Lấy giá trị số lượng
    for i in range(quantity):  # Lặp qua số lượng
        # Đảm bảo không vượt quá số lượng màu trong danh sách colors
        color_index = i % len(colors)
        
        packer.addItem(
            Item(
                partno=row['PART NO'],
                name='test' + str(i),
                typeof='cube',
                WHD=(int(row['L']), int(row['W']), int(row['H'])),
                weight=int(row['Weight(kg)']),
                level=1,
                loadbear=100,
                updown=False,
                color=colors[color_index]  # Lấy màu dựa trên index
            )
        )

# calculate packing
packer.pack(
    bigger_first=True,
    # Change distribute_items=False to compare the packing situation in multiple boxes of different capacities.
    distribute_items=True,
    fix_point=True,
    check_stable=True,
    support_surface_ratio=0.75,
    number_of_decimals=1,
)

# put order
packer.putOrder()
cnt =0


# Tạo danh sách để lưu dữ liệu
data_list = []

for idx, b in enumerate(packer.bins):
    # Tính thể tích của bin
    volume = b.width * b.height * b.depth

    bin_info = {
        "Bin Name": b.string(),
        # "Bin Width": b.width,
        # "Bin Height": b.height,
        # "Bin Depth": b.depth,
        # "Space Utilization (%)": round(sum([item.width * item.height * item.depth for item in b.items]) / volume * 100, 2),
        # "Residual Volume": volume - sum([item.width * item.height * item.depth for item in b.items]),
        # "Gravity Distribution": b.gravity
    }

    for item in b.items:
        cnt = cnt + 1
        # Lưu thông tin từng item vào danh sách
        data_list.append({
            **bin_info,  # Thêm thông tin chung của bin
            "ID": item.partno,
            "Màu hiển thị": item.color,
            "Tọa độ": item.position,
            "Kiểu xoay": item.rotation_type,
            "Dimensions (W*H*D)": f"{item.width} * {item.height} * {item.depth}",
            "Volume": float(item.width) * float(item.height) * float(item.depth),
            "Cân nặng": float(item.weight),
        })


    # draw results
    painter = Painter(b)
    fig = painter.plotBoxAndItems(
        title=b.partno,
        alpha=0.8,
        write_num=True,
        fontsize=10,
    )


print("***************************************************")

print ("Số lượng thùng đã xếp",cnt)
    # Tạo DataFrame từ danh sách dữ liệu
df = pd.DataFrame(data_list)

    # Ghi dữ liệu ra file Excel
df.to_excel("packing_results.xlsx", index=False)

print("Dữ liệu đã được ghi ra file 'packing_results.xlsx'")

stop = time.time()
print('used time : ',stop - start)

fig.show()
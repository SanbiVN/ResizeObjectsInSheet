# ResizeObjectsInSheet
 Hàm tự động chỉnh ảnh:  vị trí, kích thước, sắp xếp ngang dọc


![auto resize image excel](https://user-images.githubusercontent.com/58664571/210742432-7e32038d-a228-4a93-b59c-be56b808eeeb.gif)


Hàm bổ trợ chỉnh ảnh: vị trí, kích thước, sắp xếp ảnh một cách thuận tiện, giúp các bạn bớt khó khăn trong việc tinh chỉnh ảnh khi trong trang tính của bạn có quá nhiều đối tượng hình ảnh hoặc shape.

với Hàm ResizePic

# HƯỚNG DẪN

Hàm: ResizePic(Đối số tùy thuộc điều kiện)

Các hàm bổ trợ để cài đặt đối số (Gọi là hàm cài đặt hoặc hàm định hướng):

Hàm bổ trợ | Chức năng
----------------|----------------
rp_SIZE (width, height, [Padding])|Đặt kích thước khác kích thước ô cho ảnh
rp_Sequence_V (width, height, [Padding])|Đặt kích thước và chỉ định sắp xếp dọc
rp_Sequence_H (width, height, [Padding])|Đặt kích thước và chỉ định sắp xếp ngang
rp_InCells (Cells, width, height, [Padding])	|Tên ảnh nằm trong ô, dùng hàm này để đặt

** Tham số trong cặp dấu ngoặc vuông [ ], là tham số bổ trợ.
(Padding là thêm khoảng cách giữa hai ảnh, width là chiều rộng, height là chiều cao)

### Ví dụ 1
Canh chỉnh ảnh vừa ô

Hình ảnh tên "Picture 1", vừa với ô A1:B2 (Được gộp ô), thì gõ =ResizePic("Picture 1",A1) 
Hình ảnh tên "Picture 1","Picture 2","Picture 3", vừa với ô A1:B2, A3:B4, A5:B6 (Được gộp ô) 
thì gõ =ResizePic("Picture 1",A1,"Picture 2",A3,"Picture 3",A5) 
**Bạn cần chú ý nhập, tên ảnh đứng trước ô, nếu ô chưa gộp thì hãy gõ đầy đủ tham chiếu (như A5:B6)


Nếu muốn kích thước khác ngoài kích thước ô thì sử dụng hàm rp_SIZE để đặt:
=ResizePic("Picture 1",A1,"Picture 2",A3,"Picture 3",A5, rp_SIZE(200,150,10)) 


### Ví dụ 2
Sắp xếp nhiều ảnh Ngang hoặc Dọc
Sắp xếp hình ảnh tên "Picture 1", ...2, ….3, ...4, ...5. Bắt đầu từ vị trí ô B2, theo chiều dọc 
Với kích thước ngang 200, dọc 150 và cách nhau 10
= ResizePic(B2, rp_Sequence_V(200,150,10),"Picture 1","Picture 2","Picture 3","Picture 4","Picture 5" ) 
Nếu để ngang là 0 thì ảnh sẽ giãn theo kích thước dọc.
Ngược lại, nếu để dọc là 0 thì ảnh sẽ giãn theo kích thước ngang.

Xếp theo chiều Ngang thì cài đặt đối số sử dụng hàm rp_Sequence_H 

### Ví dụ 3
Nếu tên hình ảnh nằm trong ô A1:A100 thì hãy sử dụng [hàm cài đặt] rp_InCells
thì gõ: =ResizePic(rp_InCells (A1:A100))
Kích thước ngang và dọc sẽ sao chép lại ô tương ứng.
Nếu sử dụng hàm rp_SIZE để đặt kích thước thì tất cả ảnh được đặt theo.
Thay vì gõ cài đặt kích thước: =ResizePic(rp_InCells (A1:A100), rp_SIZE(200,150))
Thì hãy gõ gọn: =ResizePic(rp_InCells(A1:A100,200,150))
Nhiều vùng hãy gõ: =ResizePic(rp_InCells(A1:A100,200,150), rp_InCells(C1:C100))

# Sao chép mã vào dự án

Để hàm hoạt động trong dự án của bạn hãy sao chép mã vào một Module toàn cục (Không đặt [Option Private Module])

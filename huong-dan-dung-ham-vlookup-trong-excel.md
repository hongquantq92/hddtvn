Hướng dẫn dùng hàm VLOOKUP trong Excel
======================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-dung-ham-vlookup-trong-excel/)Hàm VLOOKUP trong Excel là hàm tìm kiếm giá trị theo cột và trả về phương thức hàng dọc. Đây là hàm excel khá thông dụng và tiện lợi với những người hay sử dụng Excel để thống kê, dò tìm dữ liệu theo cột chính xác. Bài viết sau đây sẽ hướng dẫn cụ …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm VLOOKUP trong Excel là hàm tìm kiếm giá trị theo cột và trả về phương thức hàng dọc. Đây là hàm excel khá thông dụng và tiện lợi với những người hay sử dụng Excel để thống kê, dò tìm dữ liệu theo cột chính xác. Bài viết sau đây sẽ hướng dẫn cụ thể các bạn cách dùng hàm VLOOKUP trong Excel.**


### 1. Cú Pháp Hàm VLOOKUP


=VLOOKUP(Lookup\_value, Table\_array, Col\_index\_ num, Range\_lookup)


Trong đó


– **Lookup\_value**: Giá trị cần dò tìm.


– **Table\_array**: Bảng giới hạn để dò tìm, bạn cần F4 để Fix cố định giá trị cho mục đích copy công thức tự động.


– **Col\_index\_num**: Số thứ tự của cột lấy dữ liệu trong bảng cần dò tìm.


– **Range\_lookup**: Là giá trị Logic (TRUE=1, FALSE=0) quyết định so chính xác hay so tương đổi với bảng giới hạn.


+ Nếu Range\_lookup = 1 (TRUE): So tương đối.  

+ Nếu Range\_lookup = 0 (FALSE): So chính xác.  

+ Nếu bỏ qua đối này thì Excel hiểu là Range\_lookup = 1


### 2. Cách sử dụng hàm VLOOKUP


#### Ví dụ. Bạn hãy tính tiền thuế theo loại hàng của các mặt hàng dưới đây:


Trong ví dụ trên, tại ô G2 ta gõ công thức: **=VLOOKUP(D2;$D$14:$E$17;2;0)**


***Trong đó:***


**Vlookup:** là hàm dùng để tìm kiếm ra thuế suất


**D2:** Giá trị là đối tượng cần tìm. Ở đây là các đối tượng: **Đá, Bột đá, Đất**


**$D$17:$F$20:** Bảng giới hạn dò tìm. Chính là D17:F20 nhưng được F4 để Fix cố định giá trị để Copy công thức xuống các ô G3=>G11


**2:** Thứ tự cột giá trị cần lấy, trong trường hợp này chính là cột Thuế suất


**0:** Trường hợp này lấy giá trị chính xác nên chọn là 0, trường hợp này ta cũng có thể chọn là 1


![](https://hddtvn.com/wp-content/uploads/2021/01/s4SmGsg.png)


#### Copy công thức xuống các ô G3=>G11 ta được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/eRnXO6L.png)


*Bài viết trên đã hướng dẫn các bạn cách dùng hàm VLOOKUP trong Excel. Hy vọng qua bài viết, các bạn đã có thể sử dụng thành thạo được hàm này. Chúc các bạn thành công!*


moreHàm VLOOKUP trong Excel là hàm tìm kiếm giá trị theo cột và trả về…


Cách sử dụng hàm CONVERT để chuyển đổi đơn vị tính trong Excel
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-convert-de-chuyen-doi-don-vi-tinh-trong-excel/)Hàm CONVERT được sử dụng trong Excel để chuyển đổi đơn vị tính từ hệ thống đo lường này sang hệ thống đo lường khác. Ví dụ như chuyển đổi đơn vị công suất, diện tích, khoảng cách, trọng lượng, thời gian… Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm CONVERT được sử dụng trong Excel để chuyển đổi đơn vị tính từ hệ thống đo lường này sang hệ thống đo lường khác. Ví dụ như chuyển đổi đơn vị công suất, diện tích, khoảng cách, trọng lượng, thời gian… Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm CONVERT trong Excel.**


### 1. Cấu trúc hàm CONVERT


Hàm CONVERT hỗ trợ chuyển đổi một số dạng đơn vị đo lường như: trọng lượng, khối lượng, khoảng cách, thời gian, áp suất, lực, năng lượng, công suất, từ tính, nhiệt độ, thể tích, diện tích, thông tin…


Cú pháp hàm: **=CONVERT(number; from\_unit; to\_unit)**


*Trong đó:*




* **number**: Đối số bắt buộc, là giá trị tính bằng **from\_units** cần chuyển đổi.

* **from\_unit**: Đối số bắt buộc, là đơn vị của **number**.

* **to\_unit**: Đối số bắt buộc, là đơn vị muốn chuyển đối tới.



Chú ý:




* Nếu kiểu dữ liệu đầu vào không chính xác -> hàm trả về giá trị lỗi #VALUE!

* Nếu đơn vị không tồn tại -> hàm trả về giá trị lỗi #N/A.

* Nếu đơn vị không hỗ trợ tiền tố nhị phân -> hàm trả về giá trị lỗi #N/A.

* Nếu form\_unit, to\_unit không thuộc cùng 1 nhóm đơn vị đo lường -> hàm trả về giá trị lỗi #N/A.

* Tên và tiền tố đơn vị có phân biệt chữ hoa và chữ thường.



### 2. Cách sử dụng hàm CONVERT


Ví dụ ta có bảng các giá trị và đơn vị tính cần chuyển đổi như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/IvqgEeY.png)


Số thứ 1: Ta cần chuyển đổi 1054 gam sang tấn. Tại ô E3 ta có công thức: **=CONVERT(B3;”g”;”ton”)**


Ta thu được kết quả: 1054 gam = 0,001161836 tấn


![](https://hddtvn.com/wp-content/uploads/2021/01/Je6BrdI.png)


Số thứ 2: Ta cần chuyển đổi 220,06 mét sang yard. Tại ô E4 ta có công thức: **=CONVERT(B4;”m”;”yd”)**


Ta thu được kết quả: 220,06 mét = 240,6605424 yard


Tương tự với các số thứ 3 là chuyển đổi từ giây sang ngày, số thứ 4 là chuyển đổi từ bit sang byte, số thứ 5 là chuyển đổi từ độ atmosphere sang độ pascal.


![](https://hddtvn.com/wp-content/uploads/2021/01/opa3U0d.png)


Với số thứ 6 là chuyển từ inch sang minimet. Do 2 đơn vị này không cùng nhóm đơn vị nên không thể chuyển đổi. Lúc này Excel sẽ báo lỗi **#N/A**.


![](https://hddtvn.com/wp-content/uploads/2021/01/s8Zk2MC.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm CONVERT để chuyển đổi đơn vị tính trong Excel. Chúc các bạn thành công!*


moreHàm CONVERT được sử dụng trong Excel để chuyển đổi đơn vị tính từ hệ…


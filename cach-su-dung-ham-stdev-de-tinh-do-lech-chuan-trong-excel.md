Cách sử dụng hàm STDEV để tính độ lệch chuẩn trong Excel
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-stdev-de-tinh-do-lech-chuan-trong-excel/)Hàm STDEV là một trong những hàm thống kê thông dụng trong Excel. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm STDEV để tính độ lệch chuẩn dựa trên mẫu. 1. Cấu trúc hàm STDEV Cú pháp hàm: STDEV(number1;[number2],…) Trong đó: Number1: Đối số bắt buộc. Đối số dạng số …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm STDEV là một trong những hàm thống kê thông dụng trong Excel. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm STDEV để tính độ lệch chuẩn dựa trên mẫu.**


### 1. Cấu trúc hàm STDEV


Cú pháp hàm: **STDEV(number1;[number2],…)**


*Trong đó:*




* **Number1**: Đối số bắt buộc. Đối số dạng số đầu tiên tương ứng với mẫu tổng thể.

* **Number2,…**: Đối số tùy chọn. Đối số dạng số từ 2 đến 255 tương ứng với mẫu tổng thể. Bạn có thể thay thế các giá trị riêng lẻ bằng 1 mảng đơn hoặc tham chiếu tới 1 mảng.



*Lưu ý:*


Độ lệch chuẩn được tính bằng cách dùng phương pháp “n-1”.


Đối số có thể là số hoặc tên, mảng hoặc tham chiếu có chứa số.


Các giá trị logic và trình bày số dạng văn bản mà bạn gõ trực tiếp vào danh sách các đối số sẽ được đếm.


Nếu đối số là mảng hay tham chiếu, chỉ các số trong mảng hay tham chiếu đó mới được tính. Các ô trống, giá trị lô-gic, văn bản hoặc giá trị lỗi trong mảng hoặc tham chiếu bị bỏ qua.


Các đối số là văn bản hay giá trị lỗi không thể chuyển đổi thành số sẽ khiến xảy ra lỗi.


STDEV dùng công thức sau đây:


![Công thức](https://hddtvn.com/wp-content/uploads/2021/01/20ef393d-e7c5-4dcc-829a-aa108436654d.gif)


trong đó x là mẫu average(number1,number2,…) và n là cỡ mẫu.


### 2. Cách sử dụng hàm STDEV để tính độ lệch chuẩn


Ví dụ: So sánh biến động doanh thu trong 3 năm của các mặt hàng sau


![](https://hddtvn.com/wp-content/uploads/2021/01/X9qhAPU.png)


Ta nhập công thức: =STDEV(C2:E2) tại ô tính độ lệch chuẩn doanh thu của Rượu vodka Lạc Hồng


![](https://hddtvn.com/wp-content/uploads/2021/01/4fFxd2D.png)


Ta thu được độ lệch chuẩn doanh thu của Rượu vodka Lạc Hồng


![](https://hddtvn.com/wp-content/uploads/2021/01/twBlzAm.png)


Sao chép công thức xuống các ô bên dưới để tính độ lệch chuẩn doanh thu của các mặt hàng còn lại


![](https://hddtvn.com/wp-content/uploads/2021/01/4bpeu9L.png)


Từ kết quả trên ta thấy doanh thu của Rượu vodka Lạc Hồng là ổn định nhất vì có độ lệch chuẩn nhỏ nhất. Doanh thu của Nước khoáng bất ổn nhất do có độ lệch chuẩn lớn nhất.


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm STDEV để tính độ lệch chuẩn trong Excel. Chúc các bạn thành công!*


moreHàm STDEV là một trong những hàm thống kê thông dụng trong Excel. Bài viết…


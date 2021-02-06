Cách sử dụng hàm INDEX để tham chiếu dữ liệu trong Excel
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-index-de-tham-chieu-du-lieu-trong-excel/)INDEX là một trong số những hàm được sử dụng nhiều nhất trong Excel. Hàm này là hàm trả về mảng, giúp lấy các giá trị tại một ô nào đó giao giữa cột và dòng. Hàm INDEX cũng rất nhạy và có thể kết hợp với các hàm khác nữa để giải quyết được …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**INDEX là một trong số những hàm được sử dụng nhiều nhất trong Excel. Hàm này là hàm trả về mảng, giúp lấy các giá trị tại một ô nào đó giao giữa cột và dòng. Hàm INDEX cũng rất nhạy và có thể kết hợp với các hàm khác nữa để giải quyết được nhiều vấn đề đặt ra. Trong bài viết dưới đây, hddtvn sẽ hướng dẫn các bạn chi tiết cách sử dụng hàm INDEX nhé.**


Hàm INDEX có 2 dạng là dạng mảng và dạng tham chiếu


### 1. Hàm INDEX dạng mảng


#### a. Cấu trúc hàm INDEX dạng mảng


Cú Pháp hàm: **=INDEX(array; row\_num; [column\_num])**


*Trong đó:*




* **array**: đối số bắt buộc, là phạm vi ô hoặc một hằng số mảng.

* **row\_num**: đối số bắt buộc, là hàng trong mảng mà từ đó trả về một giá trị.

* **column\_num**: đối số tùy chọn, là cột trong mảng mà từ đó trả về một giá trị.



Lưu ý:




* Hàm INDEX trả về giá trị của một ô được xác định bởi số hàng và cột trong mảng.

* Dùng dạng mảng nếu đối số thứ nhất của hàm INDEX là một hằng số mảng.

* Nếu mảng chỉ chứa một hàng hoặc cột, đối số **row\_num** hoặc **column\_num** tương ứng là tùy chọn.

* Nếu mảng có nhiều hàng và nhiều hơn một cột và chỉ sử dụng **row\_num** hoặc **column\_num**, chỉ mục sẽ trả về một mảng của toàn bộ hàng hoặc cột trong mảng.

* Nếu **row\_num** được bỏ qua, **column\_num** là bắt buộc.

* Nếu **column\_num** được bỏ qua, **row\_num** là bắt buộc.



#### b. Cách sử dụng hàm INDEX dạng mảng


Ví dụ ta có bảng dữ liệu như hình sau đây:


![](https://hddtvn.com/wp-content/uploads/2021/01/OKCBULT.png)


Để lấy được doanh thu 2018 của Bia lon ĐV vàng, áp dụng cấu trúc hàm như trên ta có công thức:


**=INDEX(A1:E10;62;4)**


*Trong đó:*




* **A1:E10** là phạm vi bảng dữ liệu

* **6** là số hàng của Bia lon ĐV vàng

* **4** là số cột của doanh thu 2018



![](https://hddtvn.com/wp-content/uploads/2021/01/VuZMHWF.png)


### 2. Hàm INDEX dạng tham chiếu


#### a. Cấu trúc hàm INDEX dạng tham chiếu


Cú pháp hàm: **=INDEX(reference; row\_num; [column\_num]; [area\_num])**


*Trong đó:*




* **reference**: đối số bắt buộc, là vùng tham chiếu.

* **row\_num**: đối số bắt buộc, là số hàng từ đó trả về một tham chiếu.

* **column\_num**: đối số tùy chọn, là số cột từ đó trả về một tham chiếu.

* **area\_num**: đối số tùy chọn, là số của vùng ô sẽ trả về giá trị trong reference.



*Lưu ý:*




* Nếu **Area\_num** được bỏ qua thì hàm INDEX dùng vùng 1, tùy chọn.

* Nếu bạn đặt **row\_num** hoặc **column\_num** thành 0 thì hàm INDEX trả về tham chiếu cho toàn bộ cột hoặc hàng tương ứng.

* **row\_num**, **column\_num** và **area\_num** phải tới một ô trong tham chiếu; Nếu không, chỉ mục trả về #REF!



#### b. Cách sử dụng hàm INDEX dạng tham chiếu


Cũng với ví dụ như trên, áp dụng cấu trúc hàm ta có công thức tìm doanh thu 2018 của Bia lon ĐV vàng như sau:


**=INDEX(A1:E10;6;4;1)**


*rong đó:*




* **A1:E10** là phạm vi bảng dữ liệu

* **6** là số hàng của Bia lon ĐV vàng

* **4** là số cột của doanh thu 2018

* **1** là số của vùng ô sẽ trả về giá trị



![](https://hddtvn.com/wp-content/uploads/2021/01/N5fzC4l.png)


*Như vậy, bài viết trên đã chỉ ra hai dạng tham chiếu của hàm INDEX cũng như cách sử dụng hàm. Chúc các bạn thành công!*


moreINDEX là một trong số những hàm được sử dụng nhiều nhất trong Excel. Hàm…


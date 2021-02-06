Hướng dẫn sử dụng hàm FLOOR để làm tròn số xuống trong Excel
============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-su-dung-ham-floor-de-lam-tron-so-xuong-trong-excel/)Excel hỗ trợ người dùng rất nhiều hàm làm tròn số theo nhiều yêu cầu khác nhau. Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng hàm FLOOR làm tròn số xuống tới bội số có nghĩa gần nhất trong Excel. 1. Cấu trúc hàm FLOOR Cú pháp hàm: =FLOOR(number; …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel hỗ trợ người dùng rất nhiều hàm làm tròn số theo nhiều yêu cầu khác nhau. Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng hàm FLOOR làm tròn số xuống tới bội số có nghĩa gần nhất trong Excel.**


### 1. Cấu trúc hàm FLOOR


Cú pháp hàm: **=FLOOR(number; significance)**


*Trong đó:*




* **Number**: đối số bắt buộc, là số mà bạn muốn làm tròn.

* **Significance**: đối số bắt buộc, là bội số mà bạn muốn làm tròn đến.



*Lưu ý:*




* Nếu đối số **number** dương và đối số **significance** âm. Thì hàm FLOOR trả về giá trị lỗi #NUM! .

* Nếu một trong hai đối số không phải là số thì hàm FLOOR trả về giá trị lỗi #VALUE! .

* Nếu dấu của **number** là dương, thì giá trị được làm tròn xuống và điều chỉnh tiến tới không.

* Nếu dấu của **number** là âm, thì giá trị được làm tròn xuống và điều chỉnh rời xa không.

* Nếu **number** là bội số của **significance** thì hàm sẽ không làm tròn thêm.



### 2. Cách sử dụng hàm FLOOR


Ví dụ ta cần làm tròn 11 xuống đến bội số gần nhất của 5. Áp dụng cấu trúc hàm bên trên ta có công thức:


**=FLOOR(11;5)**


![](https://hddtvn.com/wp-content/uploads/2021/01/BPENiTI.png)


Tương tự ta có công thức cho các trường hợp khác như sau:




* Làm tròn -11 xuống đến bội số gần nhất của -5: =FLOOR(-11;-5)

* Làm tròn 10 xuống đến bội số gần nhất của 5: =FLOOR(10;5)

* Làm tròn -11 xuống đến bội số gần nhất của 5: =FLOOR(-11;5)

* Làm tròn 11 xuống đến bội số gần nhất của -5: =FLOOR(11;-5)



Ta có thể thấy nếu **number** là số dương và **significance** là số âm thì hàm sẽ trả về giá trị lỗi #NUM!. Các bạn cần chú ý điều này.


![](https://hddtvn.com/wp-content/uploads/2021/01/NyEXQtW.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm FLOOR để làm tròn số xuống tới bội số có nghĩa gần nhất trong Excel. Chúc các bạn thành công!*


moreExcel hỗ trợ người dùng rất nhiều hàm làm tròn số theo nhiều yêu cầu…


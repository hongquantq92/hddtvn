Cách sử dụng hàm WEEKNUM để tính số thứ tự của tuần trong năm
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-weeknum-de-tinh-so-thu-tu-cua-tuan-trong-nam/)Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm WEEKNUM để tính số thứ tự của tuần trong năm của một ngày cụ thể trong Excel. 1. Cấu trúc hàm WEEKNUM Cú pháp hàm: =WEEKNUM(serial\_number; [return\_type]) Trong đó: Serial\_number: đối số bắt buộc, là một ngày trong tuần muốn tính số …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm WEEKNUM để tính số thứ tự của tuần trong năm của một ngày cụ thể trong Excel.**


![](https://hddtvn.com/wp-content/uploads/2021/01/2020-calendar_1017-21131.jpg)


### 1. Cấu trúc hàm WEEKNUM


Cú pháp hàm: **=WEEKNUM(serial\_number; [return\_type])**


*Trong đó:*




* **Serial\_number**: đối số bắt buộc, là một ngày trong tuần muốn tính số thứ tự của tuần đó trong năm.

* **Return\_type**: đối số tùy chọn, là một số để xác định tuần sẽ bắt đầu từ ngày nào.



*Lưu ý:*




* Có hai hệ thống được dùng cho hàm này:



– Hệ thống 1 Tuần có ngày 1 tháng 1 là tuần thứ nhất trong năm và được đánh số là tuần 1.


– Hệ thống 2 Tuần có ngày thứ Năm đầu tiên trong năm là tuần thứ nhất trong năm và được đánh số là tuần 1.




* Nếu bỏ qua đối số return\_type, hàm sẽ mặc định là 1.

* Nếu **return\_type** = 1 thì tuần bắt đầu vào chủ nhật và được đánh số là 1.

* Nếu **return\_type** = 2 thì tuần bắt đầu vào thứ 2 và được đánh số là 1.

* Nếu **return\_type** = 11 thì tuần bắt đầu vào thứ 2 và được đánh số là 1.

* Nếu **return\_type** = 12 thì tuần bắt đầu vào thứ 3 và được đánh số là 1.

* Nếu **return\_type** = 13 thì tuần bắt đầu vào thứ 4 và được đánh số là 1.

* Nếu **return\_type** = 14 thì tuần bắt đầu vào thứ 5 và được đánh số là 1.

* Nếu **return\_type** = 15 thì tuần bắt đầu vào thứ 6 và được đánh số là 1.

* Nếu **return\_type** = 16 thì tuần bắt đầu vào thứ 7 và được đánh số là 1.

* Nếu **return\_type** = 17 thì tuần bắt đầu vào chủ nhật được đánh số là 1.

* Nếu **return\_type** = 21 thì tuần bắt đầu vào thứ 2 và được đánh số là 2.

* Nếu **return\_type** không nằm trong các giá trị của nó như trên thì hàm trả về giá trị lỗi #NUM!

* Nếu **serial\_number** nằm ngoài phạm vi cơ bản của giá trị ngày tháng thì hàm trả về giá trị lỗi #VALUE!



### .2 Cách sử dụng hàm WEEKNUM


Ví dụ ta cần tìm số thứ tự của các tuần theo yêu cầu như hình sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/e1GLUeH.png)


Áp dụng cấu trúc hàm WEEKNUM như trên, ta có công thức tính số thứ tự của tuần bắt đầu vào Chủ nhật có ngày 03/02/2019 như sau:


**=WEEKNUM(A3;1)**


![](https://hddtvn.com/wp-content/uploads/2021/01/GmML4Ii.png)


Áp dụng công thức hàm tương tự ta có công thức tính các yêu cầu còn lại như sau:




* Số thứ tự của tuần bắt đầu vào Thứ 2 có ngày 30/09/2018: **=WEEKNUM(A4;2)**

* Số thứ tự của tuần bắt đầu vào Thứ 2 có ngày 11/07/2017: **=WEEKNUM(A5;11)**

* Số thứ tự của tuần bắt đầu vào Thứ 3 có ngày 23/12/2016: **=WEEKNUM(A6;12)**

* Số thứ tự của tuần bắt đầu vào Thứ 4 có ngày 05/05/2015: **=WEEKNUM(A7;13)**

* Số thứ tự của tuần bắt đầu vào Thứ 6 có ngày 14/10/2014: **=WEEKNUM(A8;15)**



![](https://hddtvn.com/wp-content/uploads/2021/01/4iVc6Gj.png)


Tiếp theo với ô A9, ta có thể thấy ngày của ô không đúng nên hàm sẽ trả về lỗi #VALUE! . Với ô A10 thì giá trị **return\_type** không đúng nên hàm trả về lỗi #NUM!.


![](https://hddtvn.com/wp-content/uploads/2021/01/bCsmVgo.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm WEEKNUM trong Excel. Chúc các bạn thành công!*


moreBài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm WEEKNUM để…


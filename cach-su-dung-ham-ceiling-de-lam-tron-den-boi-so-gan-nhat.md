Cách sử dụng hàm CEILING để làm tròn đến bội số gần nhất
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-ceiling-de-lam-tron-den-boi-so-gan-nhat/)Trong các bảng tính số liệu trên Excel, việc làm tròn số là điều cần thiết để làm gọn số liệu tính và bảng tính. Khi làm tròn số thì việc tính toán tiếp theo cũng sẽ đơn giản và dễ dàng hơn. Trên Excel, để làm tròn số bạn có thể sử dụng tới …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong các bảng tính số liệu trên Excel, việc làm tròn số là điều cần thiết để làm gọn số liệu tính và bảng tính. Khi làm tròn số thì việc tính toán tiếp theo cũng sẽ đơn giản và dễ dàng hơn. Trên Excel, để làm tròn số bạn có thể sử dụng tới hàm CEILING. Hàm này được sử dụng để làm tròn số lên đến số nguyên gần nhất hoặc bội số có nghĩa gần nhất. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm CEILING chi tiết nhất.**


### 1. Cấu trúc hàm CEILING


Cú pháp hàm: **=CEILING(number; significance)**


*Trong đó:*




* **Number**: đối số bắt buộc, là giá trị cần làm tròn.

* **Significance**: đối số bắt buộc, là số cần làm tròn đến.



*Lưu ý:*




* Nếu số âm và bội số âm, giá trị được làm tròn xuống, xa số không.

* Nếu số âm và bội số dương, giá trị được làm tròn lên tiến đến số không.

* Nếu đối số không phải dạng số, CEILING trả về giá trị lỗi #VALUE! .



### 2. Cách sử dụng hàm CEILING


Ví dụ ta cần làm tròn số theo các yêu cầu như sau:




* Làm tròn 3,6 lên đến bội số gần nhất của 5

* Làm tròn 3,6 lên đến bội số gần nhất của 1

* Làm tròn 3,6 lên đến bội số gần nhất của 0,9

* Làm tròn -3,6 lên đến bội số gần nhất của 5

* Làm tròn -3,6 lên đến bội số gần nhất của -5

* Làm tròn -3,6 lên đến bội số gần nhất của -1

* Làm tròn -3,6 lên đến bội số gần nhất của 0,9

* Làm tròn -3,6 lên đến bội số gần nhất của -0,9



![](https://hddtvn.com/wp-content/uploads/2021/01/LcGQ0ng.png)


#### Theo công thức hàm, ta có công thức làm tròn cho các trường hợp như sau:




* Làm tròn 3,6 lên đến bội số gần nhất của 5: **=CEILING(2,7;5)**

* Làm tròn 3,6 lên đến bội số gần nhất của 1: **=CEILING(3,6;1)**

* Làm tròn 3,6 lên đến bội số gần nhất của 0,9: **=CEILING(3,6;0,9)**

* Làm tròn -3,6 lên đến bội số gần nhất của 5: **=CEILING(-3,6;5)**

* Làm tròn -3,6 lên đến bội số gần nhất của -5: **=CEILING(-3,6;-5)**

* Làm tròn -3,6 lên đến bội số gần nhất của -1: **=CEILING(-3,6;-1)**

* Làm tròn -3,6 lên đến bội số gần nhất của 0,9: **=CEILING(-3,6;0,9)**

* Làm tròn -3,6 lên đến bội số gần nhất của -0,9: **=CEILING(-3,6;-0,9)**



![](https://hddtvn.com/wp-content/uploads/2021/01/tVjdXfe.png)


Như kết quả trên ta thấy:




* 3,6 được làm tròn lên đến bội số gần nhất của 5 là 5

* 3,6 được làm tròn lên đến bội số gần nhất của 1 là 4

* 3,6 được làm tròn lên đến bội số gần nhất của 0,9 chính là 3,6

* -3,6 được làm tròn lên đến bội số gần nhất của 5 là 0

* -3,6 được làm tròn lên đến bội số gần nhất của -5 là -5

* -3,6 được làm tròn lên đến bội số gần nhất của -1 là -4

* -3,6 được làm tròn lên đến bội số gần nhất của 0,9 chính là -3,6

* -3,6 được làm tròn lên đến bội số gần nhất của -0,9 chính là -3,6



*Như vậy, bài viết trên đã hướng dẫn các bạn sử dụng hàm CEILING để làm tròn tới bội số gần nhất trong Excel. Chúc các bạn thành công!*


moreTrong các bảng tính số liệu trên Excel, việc làm tròn số là điều cần…


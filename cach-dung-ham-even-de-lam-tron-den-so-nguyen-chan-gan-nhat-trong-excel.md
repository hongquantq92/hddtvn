Cách dùng hàm EVEN để làm tròn đến số nguyên chẵn gần nhất trong Excel
======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-even-de-lam-tron-den-so-nguyen-chan-gan-nhat-trong-excel/)Việc làm tròn số là điều cần thiết để làm gọn số liệu tính và bảng tính Excel. Khi làm tròn số thì việc tính toán tiếp theo cũng sẽ đơn giản và dễ dàng hơn. Excel cung cấp cho người dùng rất nhiều hàm để làm tròn số. Trong bài viết này, hddtvn sẽ …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Việc làm tròn số là điều cần thiết để làm gọn số liệu tính và bảng tính Excel. Khi làm tròn số thì việc tính toán tiếp theo cũng sẽ đơn giản và dễ dàng hơn. Excel cung cấp cho người dùng rất nhiều hàm để làm tròn số. Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn hàm EVEN để làm tròn đến số nguyên chẵn gần nhất.**


### 1. Cấu trúc hàm EVEN


Cú pháp hàm: **=EVEN(number)**


*Trong đó:* **Number** đối số bắt buộc, là giá trị cần làm tròn.


*Lưu ý:*




* Nếu **number** không phải dạng số, hàm EVEN trả về giá trị lỗi #VALUE! .

* Không xét đến dấu của đối số **number**, giá trị được làm tròn ra xa số không.

* Nếu **number** là số nguyên chẵn, thì không làm tròn.



### 2. Cách sử dụng hàm EVEN


Ví dụ ta cần làm tròn những số sau đây:




* Làm tròn 5 đến số nguyên chẵn gần nhất

* Làm tròn 3,5 đến số nguyên chẵn gần nhất

* Làm tròn 4,4 đến số nguyên chẵn gần nhất

* Làm tròn -5 đến số nguyên chẵn gần nhất

* Làm tròn -3,5 đến số nguyên chẵn gần nhất

* Làm tròn -4,4 đến số nguyên chẵn gần nhất



![](https://hddtvn.com/wp-content/uploads/2021/01/kfyXlZe.png)


Áp dụng cấu trúc hàm như trên, ta có công thức tính cho các trường hợp như sau:




* Làm tròn 5 đến số nguyên chẵn gần nhất: **=EVEN(5)**

* Làm tròn 3,5 đến số nguyên chẵn gần nhất: **=EVEN(3,5)**

* Làm tròn 4,4 đến số nguyên chẵn gần nhất: **=EVEN(4,4)**

* Làm tròn -5 đến số nguyên chẵn gần nhất: **=EVEN(-5)**

* Làm tròn -3,5 đến số nguyên chẵn gần nhất: **=EVEN(-3,5)**

* Làm tròn -4,4 đến số nguyên chẵn gần nhất: **=EVEN(-4,4)**



![](https://hddtvn.com/wp-content/uploads/2021/01/MPhUvEM.png)


Nhìn vào kết quả, ta có thể thấy hàm EVEN sẽ làm tròn số ra xa số 0.




* Hàm sẽ làm tròn số 5 thành 6 chứ không phải thành 4 dù 4 và 6 đều gần 5. Tương tự với -5 sẽ được làm tròn thành -6 chứ không phải -4.

* Với 3,5 thì gần với 4 nhất nên sẽ được làm tròn thành 4. Tương tự với -3,5 được làm tròn thành -4.

* Còn với 4,4 thì sẽ được hàm làm tròn thành 6 chứ không phải 4 dù 4,4 gần với 4 hơn. Do hàm sẽ làm tròn ra xa số 0 chứ không phải gần số 0. Tương tự với -4,4 sẽ được làm tròn thành -6.



*Hy vọng qua bài viết trên, các bạn đã hiểu được cách sử dụng hàm EVEN để làm tròn trong Excel. Chúc các bạn thành công!*


moreViệc làm tròn số là điều cần thiết để làm gọn số liệu tính và…


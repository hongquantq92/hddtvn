Cách sử dụng hàm NPV trong Excel để tính giá trị hiện tại của dự án đầu tư
==========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-npv-trong-excel-de-tinh-gia-tri-hien-tai-cua-du-an-dau-tu/)Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng hàm NPV trong Excel để tính giá trị hiện tại của một dự án đầu tư bằng cách dùng lãi suất chiết khấu và một chuỗi các khoản thanh toán và thu nhập trong tương lai.  1. Cấu trúc hàm NPV …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng hàm NPV trong Excel để tính giá trị hiện tại của một dự án đầu tư bằng cách dùng lãi suất chiết khấu và một chuỗi các khoản thanh toán và thu nhập trong tương lai.**


### 1. Cấu trúc hàm NPV


Cú pháp hàm: **=NPV(rate,value1,[value2],…)**


*Trong đó:*




* **Rate**: đối số bắt buộc, là lãi suất chiết khấu.

* **Value1, value2, … Value1:** bắt buộc, các giá trị tiếp theo là tùy chọn. 1 tới 254 đối số thể hiện các khoản thanh toán và thu nhập.



*Lưu ý:*




* **Value1**, **value2** …. phải có khoảng cách thời gian bằng nhau và xảy ra vào cuối mỗi kỳ.

* Hàm NPV sử dụng thứ tự của **value1**, **value2** …. để diễn giải thứ tự của các dòng tiền.

* Những đối số là các ô trống, giá trị logic, giá trị lỗi hoặc văn bản mà không thể chuyển thành số sẽ được bỏ qua.

* Nếu đối số là mảng hay tham chiếu, chỉ các số trong mảng hay tham chiếu đó mới được tính. Các ô trống, giá trị logic, văn bản hoặc giá trị lỗi trong mảng hoặc tham chiếu bị bỏ qua.



### 2. Cách sử dụng hàm NPV


Ví dụ ta có một dự án đầu tư có mức đầu tư ban đầu là 100 triệu. Lãi suất chiết khấu là 10%/năm. Thu nhập năm đầu tiên là 20 triệu. Năm thứ 2 đầu tư thêm 50 triệu. Thu nhập của năm thứ 2 và thứ 3 là 30 triệu. Cần tính giá trị hiện tại của dự án đầu tư này.


![](https://hddtvn.com/wp-content/uploads/2021/01/GFvnyh3.png)


Áp dụng cấu trúc hàm bên trên ta có công thức tính giá trị hiện tại của khoản đầu tư như sau:


**=NPV(B4;B3;B5;B6;B7;B8)**


![](https://hddtvn.com/wp-content/uploads/2021/01/RnELuXy.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm NPV để tính giá trị hiện tại của dự án đầu tư. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong bài viết này, Ketoan.vn sẽ giới thiệu đến các bạn cách sử dụng hàm…


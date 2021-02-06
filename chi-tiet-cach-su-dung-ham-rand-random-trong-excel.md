Chi tiết cách sử dụng hàm RAND (Random) trong Excel
===================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/chi-tiet-cach-su-dung-ham-rand-random-trong-excel/)Bạn cần lấy một giá trị ngẫu nhiên trong khoảng giá trị nào đó? Để các giá trị không bị trùng lặp và mang tính khách quan khi lấy giá trị bạn nên sử dụng hàm Rand. Hàm RAND sẽ trả về một giá trị ngẫu nhiên. Đó có thể là một số thực ngẫu nhiên lớn …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn cần lấy một giá trị ngẫu nhiên trong khoảng giá trị nào đó? Để các giá trị không bị trùng lặp và mang tính khách quan khi lấy giá trị bạn nên sử dụng hàm Rand. Hàm RAND sẽ trả về một giá trị ngẫu nhiên. Đó có thể là một số thực ngẫu nhiên lớn hơn 0 và nhỏ hơn một số bất kì nào đó hay là một chữ cái thuộc bản Alphabet …Bài viết sau đây sẽ hướng dẫn chi tiết cách sử dụng hàm RAND trong Excel.**


### 1. Cấu trúc hàm RAND


Cú pháp hàm: **=RAND()**


Cú pháp hàm RAND không có đối số nào. Giá trị trả về sẽ là các số ngẫu nhiên nằm trong khoảng từ 0 tới 1


*Lưu ý:*




* Muốn lấy giá trị ngẫu nhiên nằm trong khoảng lớn hơn bằng 0 và nhỏ hơn a, sử dụng công thức: =RAND()*a

* Muốn lấy giá trị ngẫu nhiên nằm trong khoảng lớn hơn bằng a và nhỏ hơn b, sử dụng công thức: =RAND()*(b-a)+a

* Kết quả của hàm RAND sẽ thay đổi mỗi khi bạn cập nhật hay mở lại bảng tính.



### 2. Cách sử dụng hàm RAND


Ví dụ: Sử dụng hàm RAND để lấy giá trị ngẫu nhiên theo các yêu cầu như sau:




* Số ngẫu nhiên lớn hơn hoặc bằng 0 và nhỏ hơn 1

* Số ngẫu nhiên lớn hơn hoặc bằng 0 và nhỏ hơn 100

* Số ngẫu nhiên lớn hơn hoặc bằng 70 và nhỏ hơn 164

* Số nguyên ngẫu nhiên lớn hơn hoặc bằng 0 và nhỏ hơn 100

* Số nguyên ngẫu nhiên lớn hơn hoặc bằng 70 và nhỏ hơn 164



![](https://hddtvn.com/wp-content/uploads/2021/01/3ZxRI0O.png)


Áp dụng cấu trúc hàm bên trên ta có công thức tính cho các trường hợp như sau:




* Số ngẫu nhiên lớn hơn hoặc bằng 0 và nhỏ hơn 1: **=RAND()**

* Số ngẫu nhiên lớn hơn hoặc bằng 0 và nhỏ hơn 100: **=RAND()*100**

* Số ngẫu nhiên lớn hơn hoặc bằng 70 và nhỏ hơn 164: **=RAND()*(164-70)+70**



![](https://hddtvn.com/wp-content/uploads/2021/01/d5c1A0z.png)


Đối với 2 trường hợp cần lấy số nguyên, ta có thể kết hợp hàm RAND với hàm INT để làm tròn kết quả của hàm RAND về số nguyên gần nhất.


Công thức tính trong 2 trường hợp này sẽ như sau:




* Số nguyên ngẫu nhiên lớn hơn hoặc bằng 0 và nhỏ hơn 100: **=INT(RAND()*100)**

* Số nguyên ngẫu nhiên lớn hơn hoặc bằng 70 và nhỏ hơn 164: **=INT(RAND()*(164-70)+70)**



![](https://hddtvn.com/wp-content/uploads/2021/01/w2qtalw.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm RAND trong Excel. Chúc các bạn thành công!*


moreBạn cần lấy một giá trị ngẫu nhiên trong khoảng giá trị nào đó? Để…


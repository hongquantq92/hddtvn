Cách sử dụng hàm TYPE trong Excel
=================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-type-trong-excel/)Hàm TYPE được sử dụng để trả về kiểu giá trị của một ô trong Excel. Hãy theo dõi bài viết sau để biết cách sử dụng hàm TYPE trong Excel. 1. Cấu trúc hàm TYPE Cú pháp hàm: =TYPE(value) Trong đó: Value đối số bắt buộc, là bất kỳ giá trị nào như số, văn …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm TYPE được sử dụng để trả về kiểu giá trị của một ô trong Excel. Hãy theo dõi bài viết sau để biết cách sử dụng hàm TYPE trong Excel.**


### 1. Cấu trúc hàm TYPE


Cú pháp hàm: **=TYPE(value)**


*Trong đó:* **Value** đối số bắt buộc, là bất kỳ giá trị nào như số, văn bản hoặc tham chiếu đến ô chứa giá trị.




* Nếu **value** ở dạng số thì hàm trả về giá trị 1.

* Nếu **value** ở dạng văn bản thì hàm trả về giá trị 2.

* Nếu **value** là những giá trị logic thì hàm trả về giá trị 4.

* Nếu **value** là những giá trị lỗi thì hàm trả về giá trị 16.

* Nếu **value** ở dạng mảng thì hàm trả về giá trị 64.



*Lưu ý:*




* Hàm TYPE hữu dụng nhất khi kết hợp với các hàm sử dụng nhiều kiểu dữ liệu khác nhau như hàm ARGUMENT, INPUT.

* Không thể dùng hàm TYPE trong việc xác định ô có chứa công thức hay không.



### 2. Cách sử dụng hàm TYPE


Ví dụ ta có những giá trị cần tìm kiểu định dạng như sau:




* **2020**

* **Ketoan.vn**

* **FALSE**

* **#DIV/0!**

* **1,546**



![](https://hddtvn.com/wp-content/uploads/2021/01/FipFHVK.png)


Áp dụng cấu trúc hàm như trên ta có công thức tìm định dạng của ô A3 như sau:


**=TYPE(A3)**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả như sau:




* Ô A3 có dữ liệu dạng số nên kết quả trả về là 1

* Ô A4 có dữ liệu dạng văn bản nên kết quả trả về là 2

* Ô A5 có dữ liệu dạng giá trị logic nên kết quả trả về là 4

* Ô A6 có dữ liệu dạng giá trị lỗi nên kết quả trả về là 16

* Ô A7 có dữ liệu dạng mảng nên kết quả trả về là 64



![](https://hddtvn.com/wp-content/uploads/2021/01/S68r70x.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm TYPE trong Excel. Chúc các bạn thành công!*


moreHàm TYPE được sử dụng để trả về kiểu giá trị của một ô trong…


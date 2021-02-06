Hướng dẫn chi tiết cách sử dụng hàm BASE trong Excel
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-chi-tiet-cach-su-dung-ham-base-trong-excel/)Hàm BASE là một trong những hàm toán học và lượng giác được sử dụng phổ biến trong Excel. Hàm Excel Base chuyển đổi một số thành một cơ số được cung cấp (cơ số) và trả về một biểu diễn văn bản của giá trị được tính toán. 1. Cấu trúc hàm BASE Cú …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm BASE là một trong những hàm toán học và lượng giác được sử dụng phổ biến trong Excel. Hàm Excel Base chuyển đổi một số thành một cơ số được cung cấp (cơ số) và trả về một biểu diễn văn bản của giá trị được tính toán.**


### 1. Cấu trúc hàm BASE


Cú pháp hàm: **=BASE(number; radix; [min\_length])**


*Trong đó:*


**number**: đối số bắt buộc, là số muốn chuyển đổi sang dạng văn bản. **Number** phải là số nguyên và 0 ≤ **number** < 2 mũ 53.


**radix**: đối số bắt buộc, là cơ số muốn chuyển đổi. **Radix** phải là số nguyên và 2 ≤ **radix** ≤ 36.


**min\_length**: đối số tùy chọn, là độ dài tối thiểu của chuỗi văn bản trả về. **Min\_length** phải là số nguyên và lớn hơn hoặc bằng 0.


*Lưu ý:*




* Nếu **number** không phải là giá trị số, hàm BASE sẽ trả về giá trị lỗi #VALUE!

* Nếu **number**, **radix**, **min\_length** nằm ngoài giới hạn tối đa và tối thiểu, hàm BASE sẽ trả về giá trị lỗi #NUM!

* Giá trị tối đa của đối số **min\_length** là 255.

* Bất kỳ số nào không phải là số nguyên sẽ bị cắt thành số nguyên hoặc hàm sẽ lấy phần nguyên của các giá trị đó.

* Nếu giá trị trả về của hàm ngắn hơn độ dài tối thiểu được chỉ định thì hàm sẽ thêm vào kết quả nếu kết quả này số 0 đằng trước



### 2. Cách sử dụng hàm BASE


Ví dụ ta cần chuyển đổi các số theo yêu cầu như sau:




* Chuyển đổi số 35 thành cơ số 4

* Chuyển đổi số 35,5 thành cơ số 4

* Chuyển đổi số 35,5 thành cơ số 4, với độ dài tối thiểu là 5

* Chuyển đổi số 61 thành cơ số 12

* Chuyển đổi số 61,2368 thành cơ số 12

* Chuyển đổi số 61 thành cơ số 12, với độ dài tối thiểu là 9

* Chuyển đổi số 113 thành cơ số 24

* Chuyển đổi số 113,27 thành cơ số 24

* Chuyển đổi số 113,27 thành cơ số 24, với độ dài tối thiểu là 16



![](https://hddtvn.com/wp-content/uploads/2021/01/fuJeS5u.png)


Áp dụng cấu trúc hàm như trên ta có công thức chuyển đổi cho các trường hợp như sau:




* Chuyển đổi số 35 thành cơ số 4: **=BASE(35;4)**

* Chuyển đổi số 35,5 thành cơ số 4: **=BASE(35,5;4)**

* Chuyển đổi số 35,5 thành cơ số 4, với độ dài tối thiểu là 5: **=BASE(35,5;4;5)**

* Chuyển đổi số 61 thành cơ số 12: **=BASE(61;12)**

* Chuyển đổi số 61,2368 thành cơ số 12: **=BASE(61,2368;12)**

* Chuyển đổi số 61 thành cơ số 12, với độ dài tối thiểu là 9: **=BASE(61;12;9)**

* Chuyển đổi số 113 thành cơ số 24: **=BASE(113;24)**

* Chuyển đổi số 113,27 thành cơ số 24: **=BASE(113,27;24)**

* Chuyển đổi số 113,27 thành cơ số 24, với độ dài tối thiểu là 16: **=BASE(113,27;24;16)**



Kết quả ta thu được như hình dưới:


![](https://hddtvn.com/wp-content/uploads/2021/01/paDv64y.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn sử dụng hàm BASE trong Excel. Chúc các bạn thành công!*


moreHàm BASE là một trong những hàm toán học và lượng giác được sử dụng…


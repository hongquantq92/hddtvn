Cách xóa khoảng trắng thừa trong ô Excel
========================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-xoa-khoang-trang-thua-trong-o-excel/)Trong quá trình làm việc với Excel, khi bạn chắc chắn công thức hàm VLOOKUP của bạn chính xác nhưng kết quả lại trả về toàn lỗi. Điều đó là do những khoảng trắng không cần thiết ở đầu, cuối hoặc giữa số hoặc đoạn chữ trong ô Excel. Bài viết dưới đây sẽ hướng …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình làm việc với Excel, khi bạn chắc chắn công thức hàm VLOOKUP của bạn chính xác nhưng kết quả lại trả về toàn lỗi. Điều đó là do những khoảng trắng không cần thiết ở đầu, cuối hoặc giữa số hoặc đoạn chữ trong ô Excel. Bài viết dưới đây sẽ hướng dẫn bạn cách xóa bỏ khoảng trắng thừa trong dữ liệu này, đây là một trong những kĩ thuật Excel nâng cao khi làm việc với dữ liệu.**


![](https://hddtvn.com/wp-content/uploads/2021/01/vUOaXWO.png)


### 1. So sánh dữ liệu giữa các ô để phát hiện dấu cách hoặc ký tự không thấy được


Cách đơn giản nhất để so sánh dữ liệu giữa các ô A2 và D2 có giống nhau hay không, các bạn sử dụng công thức sau:


**=A2=D2**


Nếu kết quả của phép so sánh này là True có nghĩa là dữ liệu 2 ô này giống nhau.


Còn nếu kết quả của phép so sánh này là FALSE có nghĩa là dữ liệu 2 ô này không hoàn toàn giống nhau, tuy chúng trông giống nhau. Lúc này chúng ta cần sử dụng hàm TRIM để xóa khoảng trắng thừa trong ô A2.


### 2. Cách xóa bỏ khoảng trắng thừa bằng hàm TRIM


#### Bước 1:


Bạn sử dụng hàm **TRIM** để xoá khoảng trắng trắng thừa trong chuỗi văn bản. Bao gồm các khoảng trắng phía trước, sau và giữa chuỗi.


**Cấu trúc của hàm TRIM rất đơn giản: TRIM (chuỗi văn bản)**


Có thể thay chuỗi văn bản bằng ô chứa chuỗi đó.


Ví dụ: để xoá khoảng trắng trong ô **A2**, bạn dùng công thức:


**=TRIM(A2)**


Hình dưới là kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/IQqqepp.png)


#### Bước 2:


Sau đó các bạn **copy + paste** kết quả lại vào ô A2 và dùng lại hàm **VLOOKUP**. Lúc này hàm **VLOOKUP** đã cho ra được kết quả chính xác:


![](https://hddtvn.com/wp-content/uploads/2021/01/qGCiw1u.png)


*Bài viết trên đã hướng dẫn các bạn sử dụng hàm TRIM để xóa khoảng trắng thừa trong ô Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình xử lý dữ liệu Excel. Chúc các bạn thành công!*


moreTrong quá trình làm việc với Excel, khi bạn chắc chắn công thức hàm VLOOKUP…


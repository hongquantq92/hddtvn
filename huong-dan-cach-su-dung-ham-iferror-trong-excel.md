Hướng dẫn cách sử dụng hàm IFERROR trong Excel
==============================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-su-dung-ham-iferror-trong-excel/)Trong quá trình làm việc trên Excel, nếu như bảng tính của bạn có quá nhiều công thức trả về giá trị lỗi gây khó chịu thì hãy sử dụng hàm IFERROR để sửa những lỗi này. Hãy theo dõi bài viết sau để nắm rõ cách sử dụng hàm IFERROR trong Excel. 1. Cấu …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình làm việc trên Excel, nếu như bảng tính của bạn có quá nhiều công thức trả về giá trị lỗi gây khó chịu thì hãy sử dụng hàm IFERROR để sửa những lỗi này. Hãy theo dõi bài viết sau để nắm rõ cách sử dụng hàm IFERROR trong Excel.**


### 1. Cấu trúc hàm IFERROR


Cú pháp hàm: **=IFERROR(value, value\_if\_error)**


*Trong đó:*




* **Value**: đối số bắt buộc, là đối số cần kiểm tra xem có lỗi không.

* **Value\_if\_error**: đối số bắt buộc, là giá trị trả về nếu công thức của đối số **Value** bị lỗi. Các kiểu lỗi sau đây được đánh giá: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME? hoặc #NULL!.



*Lưu ý:* Hàm IFERROR trong Excel nhận 2 giá trị (hoặc biểu thức) và kiểm tra, nếu:




* Nếu giá trị đầu tiên được cung cấp không đánh giá lỗi, kết quả trả về sẽ là giá trị này.

* Nếu giá trị đầu tiên được cung cấp đánh giá lỗi, kết quả trả về là giá trị thứ 2 được cung cấp.



### 2. Cách sử dụng hàm IFERROR


Ví dụ ta có bảng tính như hình dưới. Cột Số lượng và Tiền hàng đã được cho trước. Ta cần lấy Tiền hàng chia cho Số lượng để được Đơn giá. Lúc này sẽ có một số sản phẩm có đơn giá bị lỗi do số lượng là mẫu số = 0 nên phép tính không thể thực hiện. Trong trường hợp này, nếu không muốn hiển thị lỗi #DIV/0! thì ta có thể sử dụng hàm IFERROR để đặt giá trị trả về là 0 nếu công thức bị lỗi.


![](https://hddtvn.com/wp-content/uploads/2021/01/MJnC2mz.png)


Áp dụng cấu trúc hàm bên trên ta có công thức trong trường hợp này như sau:


**=IFERROR(F2/E2;0)**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/vceV8fH.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm IFERROR trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ cách sử dụng hàm này. Chúc các bạn thành công!*


moreTrong quá trình làm việc trên Excel, nếu như bảng tính của bạn có quá…


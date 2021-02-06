Hướng dẫn sử dụng hàm INDIRECT để tạo vùng tham chiếu trong Excel
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-su-dung-ham-indirect-de-tao-vung-tham-chieu-trong-excel/)Hàm INDIRECT được sử dụng thay đổi tham chiếu tới một ô trong một công thức mà không thay đổi chính công thức đó. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm INDIRECT trong Excel. 1. Cấu trúc hàm INDIRECT Cú pháp hàm: =INDIRECT(ref\_text; [a1]) Trong đó: Ref\_text: …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm INDIRECT được sử dụng thay đổi tham chiếu tới một ô trong một công thức mà không thay đổi chính công thức đó. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm INDIRECT trong Excel.**


### 1. Cấu trúc hàm INDIRECT


Cú pháp hàm: **=INDIRECT(ref\_text; [a1])**


*Trong đó:*




* **Ref\_text**: đối số bắt buộc, là tham chiếu tới một ô có chứa kiểu tham chiếu A1 hoặc kiểu tham chiếu R1C1.

* **A1**: đôi số tùy chọn, là một giá trị chỉ rõ kiểu tham chiếu nào được chứa trong văn bản tham chiếu ô.



*Lưu ý:*




* Nếu **ref\_text** tham chiếu tới một file Excel khác, thì file Excel đó phải đang mở. Nếu file Excel đó không mở, thì hàm INDIRECT trả về giá trị lỗi #REF! .

* Nếu **ref\_text** tham chiếu tới một phạm vi ô bên ngoài giới hạn hàng 1.048.576 hoặc giới hạn cột 16.384 thì hàm INDIRECT trả về lỗi #REF! .

* Nếu **a1** là TRUE hoặc được bỏ qua, thì văn bản tham chiếu được hiểu là tham chiếu kiểu A1.

* Nếu **a1** là FALSE, thì văn bản tham chiếu được hiểu là tham chiếu kiểu R1C1.



### 2. Cách sử dụng hàm INDIRECT


Ví dụ ta có bảng dữ liệu như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/JMQArVP.png)


Ví dụ ta đặt công thức là **=INDIRECT(A2)**. Lúc này hàm sẽ tham chiếu tới ô **A2** được kết quả là “B4”. Sau đó hàm sẽ tiếp tục tham chiếu tới ô **B4** được kết quả là “10”. Do “10” không phải là dữ liệu kiểu tham chiếu nữa nên hàm sẽ dừng lại và cho ra kết quả là 10.


![](https://hddtvn.com/wp-content/uploads/2021/01/O6QE1vO.png)


Nếu ta đặt công thức là **=INDIRECT(A5&D2)**. Lúc này hàm sẽ tham chiếu tới ô **A5** được kết quả là “C” và tham chiếu tới ô **D2** được kết quả là “4”. Sau đó hàm kết hợp 2 kết quả thành “C4” và tiếp tục tham chiếu tới ô **C4** được kết quả là “11”. Do “11” không phải là dữ liệu kiểu tham chiếu nữa nên hàm sẽ dừng lại và cho ra kết quả là 11.


![](https://hddtvn.com/wp-content/uploads/2021/01/fxXLnPC.png)


*Như vậy, bài viết trên đã giới thiệu đến các bạn cách sử dụng hàm INDIRECT trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ hơn về cách sử dụng hàm này. Chúc các bạn thành công!*


moreHàm INDIRECT được sử dụng thay đổi tham chiếu tới một ô trong một công…


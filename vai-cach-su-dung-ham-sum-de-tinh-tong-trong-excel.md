Vài cách sử dụng hàm SUM để tính tổng trong Excel
=================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/vai-cach-su-dung-ham-sum-de-tinh-tong-trong-excel/)Hàm SUM là hàm tính tổng cơ bản nhất của Excel. Hàm SUM được sử dụng vô cùng thường xuyên trong mọi trường hợp. Hãy theo dõi bài viết sau để nắm rõ cách sử dụng hàm SUM trong Excel. 1. Cấu trúc hàm SUM Cú pháp hàm: =SUM(number1,[number2],…) Trong đó: number1: đối số bắt …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm SUM là hàm tính tổng cơ bản nhất của Excel. Hàm SUM được sử dụng vô cùng thường xuyên trong mọi trường hợp. Hãy theo dõi bài viết sau để nắm rõ cách sử dụng hàm SUM trong Excel.**


### 1. Cấu trúc hàm SUM


Cú pháp hàm: **=SUM(number1,[number2],…)**


*Trong đó:*




* **number1**: đối số bắt buộc, là số đầu tiên bạn muốn thêm vào. Số đó có thể là 4, tham chiếu ô như B6, hoặc ô phạm vi như B2:B8.

* **number2-255**: đối số tùy chọn, là số thứ 2 đến thứ 255 mà bạn muốn cộng. Số đó có thể là 4, tham chiếu ô như B6, hoặc ô phạm vi như B2:B8.



### 2. Lý do bạn nên sử dụng hàm SUM


Phương pháp =1+2 hoặc =A+B – Mặc dù bạn có thể nhập =1+2+3 hoặc =A1+B1+C2 và nhận được kết quả đầy đủ, chính xác nhưng bạn nên sử dụng hàm SUM trong vì các phương pháp này dễ dẫn đến lỗi vì một vài lý do:




* **Lỗi chính tả**. Nếu bạn phải nhập nhiều phần tử có dãy số quá dài vào công thức của mình. Sẽ dễ dàng hơn nhiều nếu đặt các giá trị này vào các ô riêng lẻ, rồi dùng công thức SUM. Ngoài ra, bạn có thể định dạng các giá trị khi ở trong ô, làm cho các giá trị dễ đọc hơn nhiều so với khi ở trong công thức.

* **Lỗi #VALUE!** do tham chiếu văn bản thay vì số. Công thức của bạn có thể bị hỏng nếu có bất kỳ giá trị nào không phải dạng số (văn bản) trong các ô được tham chiếu, lúc này công thức sẽ trả về lỗi #VALUE! . Hàm SUM sẽ bỏ qua giá trị văn bản và chỉ cung cấp cho bạn tổng các giá trị số.

* **Lỗi #REF!** do xóa hàng hoặc cột. Nếu bạn xóa một hàng hoặc cột thì công thức sẽ không cập nhật để loại trừ hàng đã xóa và công thức sẽ trả về lỗi #REF! , trong khi hàm SUM sẽ tự động cập nhật.

* **Công thức sẽ không cập nhật các tham chiếu khi chèn hàng hoặc cột**. Nếu bạn chèn một hàng hay cột thì công thức sẽ không cập nhật để bao gồm hàng mới được thêm vào, trong khi hàm SUM sẽ tự động cập nhật (miễn là bạn không nằm ngoài phạm vi được tham chiếu trong công thức). Điều này đặc biệt quan trọng nếu bạn mong đợi công thức của mình được cập nhật và điều này không xảy ra, vì công thức sẽ cho bạn các kết quả không hoàn chỉnh mà bạn có thể không nắm được.



### 3. Cách sử dụng hàm SUM


Ví dụ nếu ta cần tính tổng Doanh thu năm 2017 trong bảng dưới, ta chỉ cần đặt công thức:


**=SUM(C2:C10)**


Kết quả ta sẽ thu được tổng Doanh thu năm 2017 như hình dưới. Ngoài ra các bạn cũng có thể để con trỏ chuột tại ô C11 sau đó sử dụng tổ hợp phím tắt **Alt + =** để tạo công thức tính tổng nhanh các ô phía trên ô đó.


![](https://hddtvn.com/wp-content/uploads/2021/01/SYgUD4x.png)


Bạn cũng có thể tính tổng một mảng cộng một ô bằng công thức:


**=SUM(C2:C10;D2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/zSNlxIA.png)


Hoặc tính tổng Doanh thu năm 2017 và năm 2019 bằng công thức:


**=SUM(C2:C10;E2:E10)**


![](https://hddtvn.com/wp-content/uploads/2021/01/ysS1aFU.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm SUM để tính tổng trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm SUM là hàm tính tổng cơ bản nhất của Excel. Hàm SUM được sử…


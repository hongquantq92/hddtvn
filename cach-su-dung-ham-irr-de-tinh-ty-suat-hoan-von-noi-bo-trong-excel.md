Cách sử dụng hàm IRR để tính tỷ suất hoàn vốn nội bộ trong Excel
================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-irr-de-tinh-ty-suat-hoan-von-noi-bo-trong-excel/)Hàm IRR được dùng để tính tỷ suất hoàn vốn nội bộ của một chuỗi dòng tiền nhằm đánh giá hiệu quả của phương án đầu tư hoặc dự án. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm IRR trong Excel. 1. Cấu trúc của hàm IRR Cú pháp hàm: …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm IRR được dùng để tính tỷ suất hoàn vốn nội bộ của một chuỗi dòng tiền nhằm đánh giá hiệu quả của phương án đầu tư hoặc dự án. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm IRR trong Excel.**


### 1. Cấu trúc của hàm IRR


Cú pháp hàm: **=IRR(values; [guess])**


*Trong đó:*




* **Values**: Đối số bắt buộc. Là một mảng hoặc các tham chiếu đến các ô có chứa số liệu của dòng tiền. Giá trị đầu tư ban đầu: là 1 số âm. Những giá trị tiếp theo: là lợi nhuận hàng năm của dự án. Lưu ý các giá trị này phải theo trình tự thời gian.

* **Guess**: số % ước lượng gần với kết quả của IRR, thường mặc định là 10%.



*Lưu ý:*




* Values phải chứa ít nhất 1 giá trị âm và 1 giá trị dương.

* Hàm IRR sử dụng thứ tự các giá trị của values như là thứ tự lưu động tiền mặt. Do đó cần cẩn thận để các thứ tự chi trả hoặc thu nhập luôn được nhập đúng.

* Hàm IRR chỉ tính toán các giá trị số bên trong các mảng hoặc tham chiếu của values; còn các ô rỗng, các giá trị logic, text hoặc các giá trị lỗi đều sẽ bị bỏ qua.

* Excel dùng chức năng lặp trong phép tính IRR. Bắt đầu với guess, IRR lặp cho tới khi kết quả chính xác trong khoảng 0.00001%. Nếu IRR không thể đưa ra kết quả sau 20 lần lặp, IRR sẽ trả về giá trị lỗi #NUM!

* Hàm IRR có quan hệ mật thiết với hàm NPV, vì tỷ suất mà hàm IRR trả về chính là lãi suất làm sao để NPV=0.



### 2. Cách sử dụng hàm IRR


Vi dụ ta có một dự án đầu tư có chi phí ban đầu thời điểm dự án bắt đầu đi vào hoạt động sản xuất là 100 triệu. Doanh thu, chi phí, lợi nhuận của dự án trong 5 năm như bảng phân tích dự án trong hình dưới. Yêu cầu tính IRR các năm của dự án.


![](https://hddtvn.com/wp-content/uploads/2021/01/Czd8Fbq.png)


Trong bảng phân tích trên ta thấy chi phí đầu tư ban đầu của dự án là 100.000.000 => Khi xét trong IRR thì giá trị này phải mang dấu âm, vì đó là khoản tiền phải chi ra


Lợi nhuận hàng năm của dự án là các số dương => Khi xét trong IRR thì giá trị này mang dấu dương, vì đó là khoản tiền thu về.


Công thức tính IRR của các năm như sau:




* Năm thứ 1: =IRR(G5:G6) = -73%

* Năm thứ 2: =IRR(G5:G7) = -32%

* Năm thứ 3: =IRR(G5:G8) = -8%

* Năm thứ 4: =IRR(G5:G9) = 6%

* Năm thứ 5: =IRR(G5:G10) = 15%



Như vậy sau 4 năm, dự án sẽ có tính khả thi khi IRR là số dương.


![](https://hddtvn.com/wp-content/uploads/2021/01/qAPny2A.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm IRR để tính tỷ suất hoàn vốn nội bộ trong Excel. Chúc các bạn thành công!*


moreHàm IRR được dùng để tính tỷ suất hoàn vốn nội bộ của một chuỗi…


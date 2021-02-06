Cách sử dụng hàm OFFSET để tham chiếu dữ liệu trong Excel
=========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-offset-de-tham-chieu-du-lieu-trong-excel/)Hàm OFFSET được sử dụng để tham chiếu hoặc tính toán dữ liệu của một ô hay một dãy trong bảng tính dựa vào một vùng tham chiếu có sẵn trước đó. Bài viết sau đâu sẽ hướng dẫn các bạn cách sử dụng hàm OFFSET trong Excel. 1. Cấu trúc hàm OFFSET Công thức …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm OFFSET được sử dụng để tham chiếu hoặc tính toán dữ liệu của một ô hay một dãy trong bảng tính dựa vào một vùng tham chiếu có sẵn trước đó. Bài viết sau đâu sẽ hướng dẫn các bạn cách sử dụng hàm OFFSET trong Excel.**


![](https://hddtvn.com/wp-content/uploads/2021/01/ham-offset-trong-excel.png)


### 1. Cấu trúc hàm OFFSET


Công thức hàm: **=OFFSET(reference; rows; cols; [height]; [width])**


*Trong đó:*




* **reference**: đối số bắt buộc, là vùng tham chiếu làm cơ sở cho hàm (làm điểm xuất phát) để tạo vùng tham chiếu mới.

* **rows**: đối số bắt buộc, là số dòng bên trên hoặc bên dưới **reference**, tính từ ô đầu tiên (ô ở góc trên bên trái) của **reference**. Ví dụ nếu rows là 2, sẽ có 2 dòng trả về và nằm bên dưới **reference**. Khi **rows** là số dương thì các dòng trả về nằm bên dưới **reference**, khi **rows** là số âm thì các dòng trả về nằm bên trên **reference**.

* **cols**: đối số bắt buộc, là số cột bên trái hoặc bên phải **reference**, tính từ ô đầu tiên (ô ở góc trên bên trái) của **reference**. Ví dụ nếu **cols** là 2 sẽ có 2 cột trả về và nằm bên phải của **reference**. Khi **cols** là số dương thì các cột trả về nằm bên phải **reference**, khi **cols** là số âm thì các cột trả về nằm bên trái **reference**.

* **height**: đối số tùy chọn, là số dòng của vùng tham chiếu cần trả về. Height phải là số dương.

* **width**: đối số tùy chọn, là số cột của vùng tham chiếu cần trả về. Width phải là số dương.



#### *Lưu* ý:




* Nếu các đối số **rows** và **cols** làm cho vùng tham chiếu trả về vượt khỏi phạm vi của trang tính thì hàm sẽ báo lỗi #REF.

* Nếu bỏ qua các đối số **height** và **width** thì tham chiếu trả về mặc định có cùng chiều cao và độ rộng với vùng tham chiếu.

* Nếu các đối số **height** và **width** >1 nghĩa là hàm phải trả về nhiều hơn 1 ô thì các bạn phải sử dụng công thức mảng hàm OFFSET nếu không sẽ nhận lỗi #VALUE!

* Đối số **height** và **width** phải là số dương.

* Hàm **OFFSET** không di chuyển hay làm thay đổi bất kỳ phần nào được chọn vì hàm chỉ trả về một tham chiếu. Các bạn có thể sử dụng hàm **OFFSET** kết hợp với các hàm cần đến một đối số tham chiếu.



### 2. Cách sử dụng hàm OFFSET


Ví dụ ta có bảng dữ liệu như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/5XxDbjY.png)


#### a. Tham chiếu tiền hàng của Đá mi sàng


Ta thấy tiền hàng của **Đá mi sàng** nằm tại ô F7. Nếu ta tham chiếu bắt đầu từ ô A1 thì phải di chuyển xuống 6 hàng và sang phải 5 cột.


Ta có công thức: **=OFFSET(A1;6;5)**


![](https://hddtvn.com/wp-content/uploads/2021/01/ZjAPAid.png)


Hoặc nếu ta tham chiếu bắt đầu từ ô C3 thì phải di chuyển xuống 4 hàng và sang phải 3 cột.


Ta có công thức: **=OFFSET(C3;4;3)**. Ta thu được kết quả vẫn tương tự như trên.


![](https://hddtvn.com/wp-content/uploads/2021/01/MzUnk7C.png)


#### b. Tính tổng tiền hàng của tất cả sản phẩm


Ta có thể sử dụng hàm OFFSET kết hợp cùng với hàm SUM để tính tổng phạm vi được trả về.


Ví dụ để tính tổng tiền hàng của tất cả sản phẩm. Nếu ta tham chiếu bắt đầu từ ô A1 thì phải di chuyển xuống 1 hàng và sang phải 5 dòng, sau đó lấy 10 hàng và 1 cột.


Ta có công thức: **=SUM(OFFSET(A1;1;5;10;1))**


![](https://hddtvn.com/wp-content/uploads/2021/01/JNUqww1.png)


Hoặc nếu ta tham chiếu bắt đầu từ ô C3 thì phải di chuyển lên 1 hàng và sang phải 3 dòng, sau đó lấy 10 hàng và 1 cột.


Ta có công thức: **=SUM(OFFSET(C3;-1;3;10;1))**. Kết quả thu được vẫn tương tự như tham chiếu từ ô A1.


![](https://hddtvn.com/wp-content/uploads/2021/01/tfzlebu.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm OFFSET và cách kết hợp hàm OFFSET với hàm SUM để tính tổng trong Excel. Chúc các bạn thành công!*


moreHàm OFFSET được sử dụng để tham chiếu hoặc tính toán dữ liệu của một…


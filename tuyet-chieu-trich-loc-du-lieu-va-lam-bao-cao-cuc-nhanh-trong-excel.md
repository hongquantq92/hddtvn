Tuyệt chiêu trích lọc dữ liệu và làm báo cáo cực nhanh trong Excel
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/tuyet-chieu-trich-loc-du-lieu-va-lam-bao-cao-cuc-nhanh-trong-excel/)Việc trích lọc dữ liệu và báo cáo là một yêu cầu không thể thiếu đối với kế toán. Trong Excel có khá nhiều cách để trích lọc và báo cáo dữ liệu. Tuy nhiên có một số kiểu dữ liệu nếu chỉ dùng công thức thông thường thì không thể lọc được. Lúc này, …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Việc trích lọc dữ liệu và báo cáo là một yêu cầu không thể thiếu đối với kế toán. Trong Excel có khá nhiều cách để trích lọc và báo cáo dữ liệu. Tuy nhiên có một số kiểu dữ liệu nếu chỉ dùng công thức thông thường thì không thể lọc được. Lúc này, người kế toán cần biết đến thủ thuật “công thức mảng” để giải quyết các yêu cầu phức tạp.**


### **Công thức mảng trong Excel là gì?**


Để hiểu khái niệm về loại công thức này, trước hết các bạn cần giải nghĩa được từ “mảng” trong Excel. Mảng được hiểu là một cột giá trị, một hàng giá trị hoặc kết hợp các cột và hàng giá trị với nhau.


(Ví dụ: 0,1,2,3)


Từ đó, công thức mảng được hiểu là: công thức được tạo ra bởi mảng giá trị nằm trong dấu ngoặc nhọn . Nó được excel tự động thêm vào sau khi kết thúc việc nhập công thức. Các bạn có thể thực hiện nhiều phép tính đối với một hoặc nhiều mục trong mảng. Kết quả của công thức này có thể sẽ là một kết quả duy nhất hoặc nhiều kết quả.


Nếu công thức mảng gồm nhiều ô thì được gọi là công thức đa ô. Ngược lại nếu công thức có trong một ô duy nhất thì được gọi là công thức đơn ô.


### **Cú pháp sử dụng**


Bước 1: Các bạn chọn ô hoặc vùng ô cần nhập công thức


Bước 2: Nhập công thức cần tính toán rồi nhấn tổ hợp phím Ctrl + Shift + Enter.


Thông thường để tính toán dữ liệu theo cột, nếu không sử dụng công thức mảng các bạn cần copy công thức để ra kết quả của các cột phía dưới. Còn khi sử dụng công thức mảng, chỉ cần nhấn tổ hợp phím Ctrl + Shift + Enter là toàn bộ kết quả sẽ hiện ra ngay lập tức.


Để hiểu hơn về cách sử dụng công thức mảng, các bạn hãy theo dõi những ví dụ sau:


#### Ví dụ 1:


Cho dữ liệu như bảng sau. Tính cột thành tiền thu mua sản phẩm của các nhân viên.


![Sử dụng công thức mảng để trích lọc dữ liệu và báo cáo trong Excel](https://scontent-sin6-1.xx.fbcdn.net/v/t1.15752-9/91592782_695331771237132_8032753449672114176_n.png?_nc_cat=101&_nc_sid=b96e70&_nc_ohc=dOjJKdictxIAX9PlHnB&_nc_ht=scontent-sin6-1.xx&oh=3a0912330dab90cb557ee92a6b356e4a&oe=5EA6EA67)


Thông thường tại ô F3 các bạn sẽ nhập công thức =E3*D3 enter rồi sao chép xuống các ô còn lại. Tuy nhiên khi dùng công thức được hướng dẫn trên để tính toán các bạn chỉ cần thực hiện:


Chọn vùng ô F3:F9 sau đó nhập công thức =E3:E9*D3:D9 rồi nhấn tổ hợp các phím Ctrl + Shift + Enter là có thể ra kết quả của toàn bộ các nhân viên. Như vậy sẽ tiết kiệm được rất nhiều thời gian khi tính toán.


![thành tiền thu mua sản phẩm của các nhân viên](https://scontent-sin6-2.xx.fbcdn.net/v/t1.15752-9/91425158_209074727099279_817028828399403008_n.png?_nc_cat=105&_nc_sid=b96e70&_nc_ohc=q_GmqiY-6XMAX8i4-mD&_nc_ht=scontent-sin6-2.xx&oh=b813241424381a6254ccf15e44bd89c2&oe=5EA7D0CE)


#### Ví dụ 2:


Tính tổng thành tiền thu mua sản phẩm của tất cả các nhân viên trong bảng dữ liệu trên.


Nếu như bình thường các bạn sẽ phải tính thành tiền của từng nhân viên rồi cộng lại hoặc sử dụng hàm DSUM khó hơn để tính. Nhưng với công thức này thì các bạn chỉ cần thực hiện thao tác nhập theo cú pháp:


=SUM(E3:E9*D3:D9) và ấn tổ hợp phím Ctrl + Shift + Enter là xong.


![tổng thành tiền thu mua sản phẩm của tất cả các nhân viên](https://scontent-sin6-1.xx.fbcdn.net/v/t1.15752-9/91158301_218881849180882_6254600550625574912_n.png?_nc_cat=101&_nc_sid=b96e70&_nc_ohc=ySFm1RyhDn8AX8X3Bp9&_nc_ht=scontent-sin6-1.xx&oh=65594d425fa8e30afbc9f0c28b71dae1&oe=5EA70797)


Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng công thức mảng để thay thế nhiều hàm phức tạp. Chúc các bạn sử dụng thành công!



moreViệc trích lọc dữ liệu và báo cáo là một yêu cầu không thể thiếu…


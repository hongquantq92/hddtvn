Cách tính số giờ làm việc dựa trên bảng chấm công trong Excel
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-tinh-so-gio-lam-viec-dua-tren-bang-cham-cong-trong-excel/)Trong công ty, kế toán sẽ phải quản lý giờ giấc của toàn bộ nhân viên để thực hiện công việc tính lương. Thông thường, bảng chấm công sẽ cung cấp dữ liệu giờ vào và giờ ra để tính giờ làm việc. Sau đó, bạn sẽ làm thế nào để cho ra kết quả …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong công ty, kế toán sẽ phải quản lý giờ giấc của toàn bộ nhân viên để thực hiện công việc tính lương. Thông thường, bảng chấm công sẽ cung cấp dữ liệu giờ vào và giờ ra để tính giờ làm việc. Sau đó, bạn sẽ làm thế nào để cho ra kết quả tính toán chính xác và nhanh nhất? Hãy đọc bài viết dưới đây để biết cách** **tính số giờ làm việc trong Excel (thông qua các hàm và thủ thuật) nhé.**


Ví dụ ta có bảng chấm công như hình dưới. Yêu cầu tính thời gian làm việc trong ngày của từng người.


![](https://hddtvn.com/wp-content/uploads/2021/01/sT9HIdq.png)


### 1. Tính số giờ làm việc bằng cách lấy giờ ra trừ giờ vào


Cách đơn giản nhất để tính thời gian làm việc trên bảng chấm công chính là lấy giờ ra trừ giờ vào. Tại ô D2 các bạn nhập công thức như sau:


**=C2-B2**


Sao chép công thức cho tất cả ô còn lại ta sẽ thu được kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/LiKS1q4.png)


### 2. Tính thời gian làm việc bằng cách sử dụng hàm


Bằng cách này thì ta sẽ cần sử dụng hai hàm đó là hàm HOUR và hàm MINUTE. Cấu trúc của hai hàm như sau:


#### a. Hàm HOUR


Cú pháp hàm: **=HOUR(serial\_number)**


*Trong đó:***serial\_number**: đối số bắt buộc, là thời gian có chứa giờ mà bạn muốn tách ra.


Hàm HOUR được sử dụng trong Excel để trả về giá trị là giờ trong ô Excel. Giá trị giờ được trả về ở dạng số nguyên từ 0 (12:00 SA) đến 23 (11:00 CH).


#### b. Hàm MINUTE


Cú pháp hàm: **=MINUTE(serial\_number)**


*Trong đó:***serial\_number**: đối số bắt buộc, là thời gian có chứa phút mà bạn muốn tách ra.


Hàm MINUTE được sử dụng trong Excel để trả về giá trị là phút trong ô Excel. Giá trị phút được trả về ở dạng số nguyên từ từ 0 tới 59.


Áp dụng cấu trúc của hai hàm như trên ta có công thức tính thời gian làm việc trong bảng chấm công như sau:


**=(HOUR(giờ ra)*60-HOUR(giờ vào)*60+MINUTE(giờ ra)-MINUTE(giờ vào))/60**


Tại ô D2 các bạn nhập công thức như sau:


**=(HOUR(C2)*60-HOUR(B2)*60+MINUTE(C2)-MINUTE(B2))/60**


Sao chép công thức cho các ô bên dưới. Sau đó các bạn bôi đen toàn bộ ô kết quả rồi chọn định dạng **General** tại mục **Number**.


![](https://hddtvn.com/wp-content/uploads/2021/01/uZ14O19.png)


Kết quả ta sẽ thu được thời gian làm của tất cả nhân viên như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/p50fzXG.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tính thời gian làm việc trên bảng chấm công Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong công ty, kế toán sẽ phải quản lý giờ giấc của toàn bộ nhân…


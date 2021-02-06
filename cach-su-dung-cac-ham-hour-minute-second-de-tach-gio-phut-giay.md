Cách sử dụng các hàm HOUR, MINUTE, SECOND để tách giờ, phút, giây
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-cac-ham-hour-minute-second-de-tach-gio-phut-giay/)Ngoài việc là công cụ hỗ trợ xử lý và tính toán một cách nhanh chóng,  Excel còn cung cấp các hàm xác định về thời gian (HOUR, MINUTE, SECOND) giúp các bạn xử lý các giá trị thời gian như: giờ, phút, giây một cách thuận tiện nhất. Bài viết sau đây sẽ hướng …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Ngoài việc là công cụ hỗ trợ xử lý và tính toán một cách nhanh chóng,  Excel còn cung cấp các hàm xác định về thời gian (HOUR, MINUTE, SECOND) giúp các bạn xử lý các giá trị thời gian như: giờ, phút, giây một cách thuận tiện nhất. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng các hàm HOUR, MINUTE, SECOND để tách giờ, phút, giây trong Excel.**


#### 1. Hàm HOUR


Hàm HOUR được sử dụng trong Excel để trả về giá trị là giờ trong ô Excel. Giá trị giờ được trả về ở dạng số nguyên từ 0 (12:00 SA) đến 23 (11:00 CH).


Cú pháp hàm: **=HOUR(serial\_number)**


*Trong đó:* **serial\_number**: đối số bắt buộc, là thời gian có chứa giờ mà bạn muốn tách ra.


*Lưu ý:* Giá trị thời gian là một phần của giá trị ngày và được biểu thị bằng số thập phân (ví dụ 12:00 CH được thể hiện là 0,5 vì nó là một nửa ngày).


Ví dụ ta có bảng thời gian như hình sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/R1eZ53X.png)


Để tách ra được giờ của cột Thời gian, áp dụng cấu trúc hàm bên trên ta có công thức cho ô B2 như sau:


**=HOUR(A2)**


Sao chép công thức cho các ô còn lại của cột Giờ ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/FasJjXo.png)


Ta có thể thấy tại ô A4 và A6 thì phần thời gian không xác định được nên hàm sẽ tự động coi là 12 giờ hoặc 0 giờ.


#### 2. Hàm MINUTE


Hàm MINUTE được sử dụng trong Excel để trả về giá trị là phút trong ô Excel. Giá trị phút được trả về ở dạng số nguyên từ từ 0 tới 59.


Cú pháp hàm: **=MINUTE(serial\_number)**


*Trong đó:* **serial\_number**: đối số bắt buộc, là thời gian có chứa phút mà bạn muốn tách ra.


Vẫn với ví dụ trên, để tách ra được phút của cột thời gian ta có công thức cho ô C2 như sau:


**=MINUTE(A2)**


Sao chép công thức cho các ô còn lại của cột Phút ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/DtphgUd.png)


Ta có thể thấy tại ô A4 và A6 thì phần thời gian không xác định được nên hàm sẽ tự động coi là 0.


#### 3. Hàm SECOND


Hàm SECOND được sử dụng trong Excel để trả về giá trị là giây trong ô Excel. Giá trị giây được trả về ở dạng số nguyên từ từ 0 tới 59.


Cú pháp hàm: **=SECOND(serial\_number)**


*Trong đó:* **serial\_number**: đối số bắt buộc, là thời gian có chứa giây mà bạn muốn tách ra.


Vẫn với ví dụ trên, để tách ra được giây của cột thời gian ta có công thức cho ô D2 như sau:


**=SECOND(A2)**


Sao chép công thức cho các ô còn lại của cột Giây ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/F9U4TQp.png)


Ta có thể thấy tại ô A4 và A6 thì phần thời gian không xác định được nên hàm sẽ tự động coi là 0.


Tại ô A3 không có giá trị giây nên hàm cũng sẽ tự động coi là 0.


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm để tách giờ, phút, giây trong Excel. Chúc các bạn thành công!*


moreNgoài việc là công cụ hỗ trợ xử lý và tính toán một cách nhanh…


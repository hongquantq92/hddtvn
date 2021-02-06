Hướng dẫn cách sử dụng hàm TIME để trả về giá trị thời gian trong Excel
=======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-su-dung-ham-time-de-tra-ve-gia-tri-thoi-gian-trong-excel/)Hàm TIME trả về kết quả là số thập phân cho một giá trị thời gian cụ thể. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm TIME trong Excel. 1. Cấu trúc hàm TIME Cú pháp hàm: =TIME(hour; minute; second) Trong đó: Hour: đối số bắt buộc, là …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm TIME trả về kết quả là số thập phân cho một giá trị thời gian cụ thể. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm TIME trong Excel.**


#### 1. Cấu trúc hàm TIME


Cú pháp hàm: **=TIME(hour; minute; second)**


*Trong đó:*




* **Hour**: đối số bắt buộc, là một số từ 0 đến 32767 thể hiện giờ.

* **Minute**: đối số bắt buộc, là một số từ 0 đến 32767 thể hiện phút.

* **Second**: đối số bắt buộc, là một số từ 0 đến 32767 thể hiện giây.



*Lưu ý:*




* Mọi giá trị của **hour** lớn hơn 23 sẽ được chia cho 24 và phần còn lại sẽ được coi là giá trị giờ. Ví dụ, TIME(31;0;0) = TIME(7;0;0) = 0,291666667 hoặc 7:00 SA.

* Mọi giá trị của **minute** lớn hơn 59 sẽ được chuyển đổi thành giờ và phút. Ví dụ, TIME(0,83,0) = TIME(1,23,0) = 0,057638889 hoặc 1:23 SA.

* Mọi giá trị của **second** lớn hơn 59 sẽ được chuyển đổi thành giờ, phút và giây. Ví dụ, TIME(0,0,1120) = TIME(0;18;40) = 0,012962963 hoặc 12:18:40 SA.

* Số thập phân mà hàm **TIME** trả về là một giá trị từ 0 đến 0,99988426. Thể hiện thời gian từ 0:00:00 (12:00:00 SA) đến 23:59:59 (11:59:59 CH).



#### 2. Cách sử dụng hàm TIME


Ví dụ ta cần tìm phần thập phân của thời gian 4 giờ, 23 phút, 11 giây. Áp dụng cấu trúc hàm như trên ta có công thức:


**=TIME(4;23;11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/m6EBKC1.png)


Vẫn với công thức trên, nhưng nếu định dạng ô là General trước khi nhập vào thì kết quả sẽ trả về kết quả dạng thời gian.


![](https://hddtvn.com/wp-content/uploads/2021/01/iq2uzIo.png)


Nếu giờ lớn hơn 23, phút lớn hơn 59, giây lớn hơn 59 thì thời gian sẽ được chuyển đổi thành giờ, phút, giây tương ứng.


![](https://hddtvn.com/wp-content/uploads/2021/01/5OTqHBt.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm TIME trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ được cách sử dụng hàm này. Chúc các bạn thành công!*


moreHàm TIME trả về kết quả là số thập phân cho một giá trị thời…


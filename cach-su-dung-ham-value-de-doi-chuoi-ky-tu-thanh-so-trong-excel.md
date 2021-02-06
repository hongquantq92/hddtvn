Cách sử dụng hàm VALUE để đổi chuỗi ký tự thành số trong Excel
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-value-de-doi-chuoi-ky-tu-thanh-so-trong-excel/)Hàm VALUE thường xuyên được sử dụng trong Excel để đổi chuỗi ký tự thành số. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm VALUE trong Excel. 1. Cấu trúc hàm VALUE Cú pháp hàm: =VALUE(text) Trong đó: Text: đối số bắt buộc, là văn bản được đặt …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm VALUE thường xuyên được sử dụng trong Excel để đổi chuỗi ký tự thành số. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm VALUE trong Excel.**


### 1. Cấu trúc hàm VALUE


Cú pháp hàm: **=VALUE(text)**


*Trong đó:*




* **Text**: đối số bắt buộc, là văn bản được đặt trong ngoặc kép hoặc một tham chiếu đến một ô có chứa văn bản bạn muốn chuyển đổi.



*Lưu ý:*




* Text có thể là bất kỳ định dạng ngày, thời gian hoặc số hằng định nào, miễn là được Microsoft Excel công nhận. Nếu text không phải là một trong các định dạng này, hàm VALUE trả về giá trị lỗi #VALUE! .

* Thông thường bạn không cần phải dùng hàm VALUE trong công thức vì Excel sẽ tự động chuyển đổi văn bản thành số khi cần thiết. Hàm này được cung cấp để tương thích với các chương trình bảng tính khác.



### 2. Cách sử dụng hàm VALUE


Ví dụ ta có bảng tính như hình dưới. Yêu cầu lấy 3 ký tự cuối của Mã hàng và chuyển đổi thành số.


![](https://hddtvn.com/wp-content/uploads/2021/01/SMrpHoX.png)


Trong trường hợp này, để lấy 3 ký cuối của Mã hàng ta có thể sử dụng hàm RIGHT. Ta có công thức lấy 3 số cuối của Mã hàng như sau:


**=RIGHT(B2;3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/gwpAERt.png)


Ta có thể thấy hàm RIGHT sẽ cho ra kết quả dạng text vì định dạng ban đầu của Mã hàng là text. Để chuyển đổi kết quả thành số ta cần kết hợp thêm hàm VALUE. Ta có công thức như sau:


**=VALUE(RIGHT(B2;3))**


Sao chép công thức cho các ô còn lại ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/SOGtV3S.png)


*Như vậy, bài viết trên đã giới thiệu đến các bạn cách sử dụng hàm VALUE trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ được cách sử dụng hàm này. Chúc các bạn thành công!*


moreHàm VALUE thường xuyên được sử dụng trong Excel để đổi chuỗi ký tự thành…


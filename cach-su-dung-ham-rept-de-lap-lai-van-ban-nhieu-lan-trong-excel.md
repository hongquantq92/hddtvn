Cách sử dụng hàm REPT để lặp lại văn bản nhiều lần trong Excel
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-rept-de-lap-lai-van-ban-nhieu-lan-trong-excel/)Trong bài viết này, hddtvn sẽ hướng dẫn các bạn cách sử dụng hàm REPT để lặp lại văn bản nhiều lần và cách sử dụng hàm REPT để vẽ biểu đồ trong Excel. 1. Cấu trúc hàm REPT Cú pháp hàm: =REPT(text, number\_times) Trong đó: Text: đối số bắt buộc, là văn bản mà …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn sẽ hướng dẫn các bạn cách sử dụng hàm REPT để lặp lại văn bản nhiều lần và cách sử dụng hàm REPT để vẽ biểu đồ trong Excel.**


### 1. Cấu trúc hàm REPT


Cú pháp hàm: **=REPT(text, number\_times)**


*Trong đó:*




* **Text**: đối số bắt buộc, là văn bản mà bạn muốn lặp lại.

* **Number\_times**: đối số bắt buộc, là số lần lặp lại văn bản.



*Lưu ý:*




* Nếu **number\_times** bằng 0 thì hàm REPT trả về văn bản trống.

* Nếu **number\_times** không phải là số nguyên thì nó bị cắt phần thập phân.

* Nếu **number\_times** là số âm hoặc không phải là số thì hàm trả về lỗi #VALUE!.

* Kết quả của hàm REPT không được dài quá 32.767 ký tự. Nếu dài hơn thì hàm REPT trả về lỗi #VALUE!.



### 2. Cách sử dụng hàm REPT


Ví dụ ta có bảng dữ liệu như hình dưới. Để lặp lại mã hàng 3 lần ta có công thức như sau:


**=REPT(A2;3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/uSzI198.png)


Trong trường hợp đối số **number\_times** là số âm thì hàm sẽ trả về giá trị lỗi #VALUE!.


**=REPT(A3;-3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/cZDNnGI.png)


Một ứng dụng khá hay khác của hàm REPT là vẽ biểu đồ. Đầu tiên các bạn chọn font chữ là Stencil. Tiếp theo các bạn chọn ký tự lặp là “|”, số lần lặp là bằng tỷ lệ nhân với 100. Công thức ta có như sau:


**=REPT(“|”;D2*100)**


Sao chép công thức cho các hàng bên dưới rồi chọn màu chữ nổi bật là ta đã có biểu đồ biểu diễn tỷ trọng các mặt hàng như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/fxIhEDc.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm REPT trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn hành công!*


moreTrong bài viết này, Ketoan.vn sẽ hướng dẫn các bạn cách sử dụng hàm REPT…


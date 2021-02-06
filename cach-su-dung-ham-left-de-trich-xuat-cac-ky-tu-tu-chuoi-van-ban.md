Cách sử dụng hàm LEFT để trích xuất các ký tự từ chuỗi văn bản
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-left-de-trich-xuat-cac-ky-tu-tu-chuoi-van-ban/)Hàm LEFT là một trong những hàm cơ bản của Excel giúp người dùng trích xuất các ký tự tính từ bên trái chuỗi văn bản. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm LEFT trong Excel. 1. Cấu trúc hàm LEFT Cú pháp hàm: =LEFT(text,[num\_chars]) Trong đó: Text: …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm LEFT là một trong những hàm cơ bản của Excel giúp người dùng trích xuất các ký tự tính từ bên trái chuỗi văn bản. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm LEFT trong Excel.**


### 1. Cấu trúc hàm LEFT


Cú pháp hàm: **=LEFT(text,[num\_chars])**


*Trong đó:*




* **Text**: đối số bắt buộc, là chuỗi văn bản có chứa các ký tự mà bạn muốn trích xuất.

* **Num\_chars**: đối số tùy chọn, là số ký tự mà bạn muốn hàm LEFT trích xuất.



*Lưu ý:*




* Hàm LEFT sẽ trích xuất số ký tự tính từ ký tự đầu tiên bên trái của văn bản.

* **Num\_chars** phải lớn hơn hoặc bằng không.

* Nếu **num\_chars** lớn hơn độ dài của văn bản, hàm LEFT trả về toàn bộ văn bản.

* Nếu **num\_chars** được bỏ qua, thì nó được mặc định là 1.

* Nếu **num\_chars** là số âm, hàm sẽ trả về giá trị lỗi #VALUE!



### 2. Cách sử dụng hàm LEFT


Ví dụ ta có bảng dữ liệu như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/3yE1ZH6.png)


Để trích xuất được 3 ký tự đầu tiên bên trái của Mã hàng, áp dụng cấu trúc hàm LEFT như trên ta có công thức như sau:


**=LEFT(A2;3)**


Sao chép công thức cho các ô còn lại ta thu được kết quả như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/6NTN7yN.png)


Nếu cần trích xuất ký tự đầu tiên của Mã hàng ta có công thức như sau:


**=LEFT(A2)**


Lúc này ta có thể bỏ qua đối số **num\_chars** do chỉ cần trích xuất ký tự đầu tiên. Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/Oae4309.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm LEFT trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ cách sử dụng hàm này. Chúc các bạn thành công!*


#### 


moreHàm LEFT là một trong những hàm cơ bản của Excel giúp người dùng trích…


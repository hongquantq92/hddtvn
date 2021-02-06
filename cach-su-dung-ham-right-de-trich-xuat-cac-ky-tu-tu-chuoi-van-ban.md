Cách sử dụng hàm RIGHT để trích xuất các ký tự từ chuỗi văn bản
===============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-right-de-trich-xuat-cac-ky-tu-tu-chuoi-van-ban/)Hàm RIGHT là một trong những hàm cơ bản của Excel giúp người dùng trích xuất các ký tự từ chuỗi văn bản. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm RIGHT trong Excel. 1. Cấu trúc hàm RIGHT Cú pháp hàm: =RIGHT(text,[num\_chars]) Trong đó: Text: đối số …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm RIGHT là một trong những hàm cơ bản của Excel giúp người dùng trích xuất các ký tự từ chuỗi văn bản. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm RIGHT trong Excel.**


### 1. Cấu trúc hàm RIGHT


Cú pháp hàm: **=RIGHT(text,[num\_chars])**


*Trong đó:*




* **Text**: đối số bắt buộc, là chuỗi văn bản có chứa các ký tự mà bạn muốn trích xuất.

* **Num\_chars**: đối số tùy chọn, là số ký tự mà bạn muốn hàm RIGHT trích xuất.



*Lưu ý:*




* **Num\_chars** phải lớn hơn hoặc bằng không.

* Nếu **num\_chars** lớn hơn độ dài của văn bản, hàm RIGHT trả về toàn bộ văn bản.

* Nếu **num\_chars** được bỏ qua, thì nó được mặc định là 1.

* Nếu **num\_chars** là số âm, hàm sẽ trả về giá trị lỗi #VALUE!



### 2. Cách sử dụng hàm RIGHT


Ví dụ ta có bảng dữ liệu như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/KarLGeg.png)


Để trích xuất được 3 ký tự cuối cùng bên phải của Mã hàng, áp dụng cấu trúc hàm RIGHT như trên ta có công thức như sau:


**=RIGHT(B2;3)**


Sao chép công thức cho các ô còn lại ta thu được kết quả như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/crqm2J7.png)


Tương tự để trích xuất ký tự cuối cùng bên phải của Mã hàng ta có công thức như sau:


**=RIGHT(B2)**


Lúc này ta có thể bỏ qua đối số **num\_chars** do chỉ cần trích xuất ký tự cuối cùng. Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/z4c6ehr.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm RIGHT trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ cách sử dụng hàm này. Chúc các bạn thành công!*


moreHàm RIGHT là một trong những hàm cơ bản của Excel giúp người dùng trích…


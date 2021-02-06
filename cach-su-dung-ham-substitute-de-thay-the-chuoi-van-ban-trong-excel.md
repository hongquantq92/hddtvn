Cách sử dụng hàm SUBSTITUTE để thay thế chuỗi văn bản trong Excel
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-substitute-de-thay-the-chuoi-van-ban-trong-excel/)Hàm SUBSTITUTE được sử dụng trong Excel để thay thế chuỗi văn bản cũ thành chuỗi văn bản mới. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm SUBSTITUTE trong Excel. 1. Cấu trúc hàm SUBSTITUTE Cú pháp hàm: =SUBSTITUTE(text, old\_text, new\_text, [instance\_num]) Trong đó: Text: đối số bắt …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm SUBSTITUTE được sử dụng trong Excel để thay thế chuỗi văn bản cũ thành chuỗi văn bản mới. Hãy theo dõi bài viết sau để hiểu rõ hơn về cách sử dụng hàm SUBSTITUTE trong Excel.**


### 1. Cấu trúc hàm SUBSTITUTE


Cú pháp hàm: **=SUBSTITUTE(text, old\_text, new\_text, [instance\_num])**


*Trong đó:*




* **Text**: đối số bắt buộc, là văn bản hoặc tham chiếu đến ô chứa văn bản mà bạn muốn thay thế các ký tự trong đó.

* **Old\_text**: đối số bắt buộc, là văn bản mà bạn muốn được thay thế.

* **New\_text**: đối số bắt buộc, là văn bản mà bạn muốn thay thế cho **old\_text**.

* **Instance\_num**: đối số tùy chọn, xác định số lần xuất hiện của **old\_text** mà bạn muốn thay bằng **new\_text**. Nếu bạn xác định **instance\_num**, thì số lần xuất hiện đó của **old\_text** được thay thế. Nếu không, mọi lần xuất hiện của **old\_text** trong văn bản được đổi thành **new\_text**.



### 2. Cách sử dụng hàm SUBSTITUTE


Ví dụ ta có bảng dữ liệu như hình dưới. Cần sửa lại Mã hàng theo các yêu cầu dưới đây.


![](https://hddtvn.com/wp-content/uploads/2021/01/U0xGqfj.png)


#### a. Đổi chữ VT thành HH


Để thực hiện đổi chữ VT thành HH, ta có công thức cho mã hàng đầu tiên như sau:


**=SUBSTITUTE(B2;”VT”;”HH”)**


Sao chéo công thức cho các ô còn lại ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/SZye6Sm.png)


#### b. Đổi tất cả số 0 trong mã thành số 9


Để thực hiện đổi tất cả số 0 trong mã thành số 9, ta có công thức cho mã hàng đầu tiên như sau:


**=SUBSTITUTE(B2;0;9)**


Sao chéo công thức cho các ô còn lại ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/lTfyczR.png)


#### c. Đổi số 0 đầu tiên trong mã thành số 9


Để thực hiện đổi số 0 đầu tiên trong mã thành số 9, ta có công thức cho mã hàng đầu tiên như sau:


**=SUBSTITUTE(B2;0;9;1)**


Sao chéo công thức cho các ô còn lại ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/HFfUeep.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm SUBSTITUTE trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm SUBSTITUTE được sử dụng trong Excel để thay thế chuỗi văn bản cũ thành…


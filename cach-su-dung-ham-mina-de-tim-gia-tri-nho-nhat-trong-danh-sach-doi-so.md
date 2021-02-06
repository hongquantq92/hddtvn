Cách sử dụng hàm MINA để tìm giá trị nhỏ nhất trong danh sách đối số
====================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-mina-de-tim-gia-tri-nho-nhat-trong-danh-sach-doi-so/)Hàm MINA được sử dụng trong Excel để tìm giá trị nhỏ nhất trong một danh sách đối số. Hãy theo dõi bài viết sau để biết cách sử dụng hàm MINA trong Excel. 1. Cấu trúc hàm MINA Cú pháp hàm: =MINA(value1,[value2],…) Trong đó: Value1: đối số bắt buộc, là đối số dạng số thứ …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm MINA được sử dụng trong Excel để tìm giá trị nhỏ nhất trong một danh sách đối số. Hãy theo dõi bài viết sau để biết cách sử dụng hàm MINA trong Excel.**


### 1. Cấu trúc hàm MINA


Cú pháp hàm: **=MINA(value1,[value2],…)**


*Trong đó:*




* **Value1**: đối số bắt buộc, là đối số dạng số thứ nhất mà bạn muốn tìm giá trị nhỏ nhất trong đó.

* **Value2,…**: đối số tùy chọn, là các đối số dạng số thứ 2 đến 255 mà bạn muốn tìm giá trị nhỏ nhất trong đó.



*Lưu ý:*




* Đối số có thể là: số, tên, mảng hay tham chiếu có chứa số; biểu thị số dạng văn bản… Ví dụ TRUE và FALSE, trong một tham chiếu.

* Các giá trị trình bày số dạng văn bản mà bạn gõ trực tiếp vào danh sách các đối số sẽ được đếm.

* Nếu đối số là mảng hay tham chiếu, chỉ các giá trị trong mảng hay tham chiếu đó mới được dùng. Các giá trị văn bản và ô trống trong mảng hay tham chiếu sẽ bị bỏ qua.

* Các đối số là văn bản hay giá trị lỗi không thể chuyển đổi thành số sẽ khiến xảy ra lỗi.

* Đối số chứa TRUE được coi là 1; đối số chứa văn bản hoặc FALSE được coi là 0.

* Nếu các đối số không chứa giá trị, hàm MINA trả về 0.

* Nếu bạn không muốn bao gồm các giá trị và dạng biểu thị số bằng văn bản trong một tham chiếu như là một phần của tính toán, hãy dùng hàm MIN.



### 2. Cách sử dụng hàm MINA


Ví dụ ta có danh sách giá trị cần tìm giá trị nhỏ nhất như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/pvQzm6E.png)


Để tìm giá trị nhỏ nhất của danh sách trên ta có công thức như sau:


**=MINA(A2:F2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/Q7zqGjX.png)


Còn trong trường hợp dưới đây thì giá trị lớn nhất sẽ là 1 do ô F2 chứa đối số TRUE được coi là 1 và là giá trị nhỏ nhất trong danh sách.


![](https://hddtvn.com/wp-content/uploads/2021/01/F1aXQSq.png)


Còn trong trường hợp sau thì giá trị lớn nhất sẽ là 0 do ô A2 chứa đối số là văn bản được coi là 0 và là giá trị nhỏ nhất trong danh sách.


![](https://hddtvn.com/wp-content/uploads/2021/01/W1yaFQD.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm MINA trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ cách sử dụng hàm này. Chúc các bạn thành công!*


#### 


moreHàm MINA được sử dụng trong Excel để tìm giá trị nhỏ nhất trong một…


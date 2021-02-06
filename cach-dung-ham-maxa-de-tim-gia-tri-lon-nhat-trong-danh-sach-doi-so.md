Cách dùng hàm MAXA để tìm giá trị lớn nhất trong danh sách đối số
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-maxa-de-tim-gia-tri-lon-nhat-trong-danh-sach-doi-so/)Hàm MAXA được sử dụng trong Excel để tìm giá trị lớn nhất trong một danh sách đối số. Hãy theo dõi bài viết sau để biết cách sử dụng hàm MAXA trong Excel. 1. Cấu trúc hàm MAXA Cú pháp hàm: =MAXA(value1,[value2],…) Trong đó: Value1: đối số bắt buộc, là đối số dạng số …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm MAXA được sử dụng trong Excel để tìm giá trị lớn nhất trong một danh sách đối số. Hãy theo dõi bài viết sau để biết cách sử dụng hàm MAXA trong Excel.**


### 1. Cấu trúc hàm MAXA


Cú pháp hàm: **=MAXA(value1,[value2],…)**


*Trong đó:*




* **Value1**: đối số bắt buộc, là đối số dạng số thứ nhất mà bạn muốn tìm giá trị lớn nhất trong đó.

* **Value2,…**: đối số tùy chọn, là các đối số dạng số thứ 2 đến 255 mà bạn muốn tìm giá trị lớn nhất trong đó.



*Lưu ý:*




* Đối số có thể là: số, tên, mảng hay tham chiếu có chứa số; biểu thị số dạng văn bản… Ví dụ TRUE và FALSE, trong một tham chiếu.

* Các giá trị trình bày số dạng văn bản mà bạn gõ trực tiếp vào danh sách các đối số sẽ được đếm.

* Nếu đối số là mảng hay tham chiếu, chỉ các giá trị trong mảng hay tham chiếu đó mới được dùng. Các giá trị văn bản và ô trống trong mảng hay tham chiếu sẽ bị bỏ qua.

* Các đối số là văn bản hay giá trị lỗi không thể chuyển đổi thành số sẽ khiến xảy ra lỗi.

* Đối số chứa TRUE được coi là 1; đối số chứa văn bản hoặc FALSE được coi là 0.

* Nếu các đối số không chứa giá trị, hàm MAXA trả về 0.

* Nếu bạn không muốn bao gồm các giá trị và dạng biểu thị số bằng văn bản trong một tham chiếu như là một phần của tính toán, hãy dùng hàm MAX.



### 2. Cách sử dụng hàm MAXA


Ví dụ ta có danh sách giá trị cần tìm giá trị lớn nhất như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/7acUJV2.png)


Để tìm giá trị lớn nhất của danh sách trên ta có công thức như sau:


**=MAXA(A2:F2)**


Kết quả ta thu được sẽ là 1 do TRUE được coi là 1 và là lớn nhất trong danh sách trên.


![](https://hddtvn.com/wp-content/uploads/2021/01/HmUJeCk.png)


Còn trong trường hợp dưới đây thì giá trị lớn nhất sẽ là 0 do ô F2 chứa văn bản được coi là 0 và là giá trị lớn nhất trong danh sách.


![](https://hddtvn.com/wp-content/uploads/2021/01/wjiVfr9.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm MAXA trong Excel. Hy vọng qua bài viết này, các bạn đã nắm rõ cách sử dụng hàm này. Chúc các bạn thành công!*


moreHàm MAXA được sử dụng trong Excel để tìm giá trị lớn nhất trong một…


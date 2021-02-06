Cách sử dụng hàm DMIN để tìm giá trị nhỏ nhất thỏa mãn điều kiện cho trước
==========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-dmin-de-tim-gia-tri-nho-nhat-thoa-man-dieu-kien-cho-truoc/)Nếu Hàm MIN dùng để tìm giá trị nhỏ nhất của các giá trị thì hàm DMIN sẽ trả về giá trị nhỏ nhất theo điều kiện các bạn chỉ định. Khi cần lấy giá trị nhỏ nhất theo một điều kiện nào đó thì các bạn nên sử dụng hàm DMIN. Ví dụ như …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Nếu Hàm MIN dùng để tìm giá trị nhỏ nhất của các giá trị thì hàm DMIN sẽ trả về giá trị nhỏ nhất theo điều kiện các bạn chỉ định. Khi cần lấy giá trị nhỏ nhất theo một điều kiện nào đó thì các bạn nên sử dụng hàm DMIN. Ví dụ như tìm chi phí mua/tiền lãi/giá bán… nhỏ nhất (rẻ nhất) của mặt hàng trong danh sách các sản phẩm. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm DMIN trong Excel.**


![Cách sử dụng hàm DMIN để tìm giá trị nhỏ nhất thỏa mãn điều kiện cho trước](https://hddtvn.com/wp-content/uploads/2021/01/dmin.png "Cách sử dụng hàm DMIN để tìm giá trị nhỏ nhất thỏa mãn điều kiện cho trước")


### 1. Cấu trúc hàm DMIN


Cú pháp hàm: **=DMIN(database; field; criteria)**


*Trong đó:*




* **database**: đối số bắt buộc, danh sách hay cơ sở dữ liệu có liên quan bao gồm cả các Tiêu đề cột.

* **field**: đối số bắt buộc, trường (cột) cần lấy giá trị nhỏ nhất. Các bạn có thể nhập trực tiếp tiêu đề cột trong dấu ngoặc kép hay một số thể hiện vị trí cột trong **database** hoặc cũng có thể nhập ô chứa tiêu đề cột cần sử dụng.

* **criteria**: đối số bắt buộc, là phạm vi ô chứa điều kiện. Các bạn có thể chọn bất kỳ miễn sao phạm vi đó chứa ít nhất một Tiêu đề cột và ô dưới tiêu đề cột chứa điều kiện cho cột.



*Lưu ý:*




* Nên đặt phạm vi điều kiện criteria trên trang tính để khi thêm dữ liệu thì phạm vi chứa điều kiện không thay đổi.

* Mặc dù phạm vi điều kiện có thể được định vị ở bất kỳ đâu trên trang tính. Nhưng phạm vi điều kiện cần tách rời không chèn lên danh sách hay cơ sở dữ liệu cần xử lý.

* **criteria** bắt buộc phải chứa ít nhất Tiêu đề cột và một ô chứa điều kiện dưới tiêu đề cột.



### 2. Cách sử dụng hàm DMIN


Ví dụ ta có bảng dữ liệu như bảng dưới và yêu cầu tìm tiền hàng nhỏ nhất của loại hàng là Bột đá


![](https://hddtvn.com/wp-content/uploads/2021/01/OMEg0eD.png)


Đầu tiên ta cần tạo điều kiện loại hàng là Bột đá


![](https://hddtvn.com/wp-content/uploads/2021/01/sj8OKLu.png)


Tiếp theo ta áp dụng công thức DMIN:


**=DMIN(A1:F11;F1;D14:D15)**


*Trong đó:*




* **A1:F11** là bảng dữ liệu

* **F1** là cột TIỀN HÀNG, cột cần tìm giá trị cao nhất.

* **D14:D15** là vùng điều kiện criteria.



Kết quả ta sẽ thu được tiền hàng nhỏ nhất của Bột đá như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/MWy2Vkh.png)


Hoặc cũng với công thức trên, ta có thể thay **field** thành số 6 là vị trí của cột TIỀN HÀNG trong bảng.


**=DMIN(A1:F11;6;D14:D15)**


Ta cũng thu được kết quả tương tự như trên


![Cách dùng hàm DMIN để tìm giá trị nhỏ nhất thỏa mãn điều kiện cho trước](https://hddtvn.com/wp-content/uploads/2021/01/1U4laEy.png "Cách dùng hàm DMIN để tìm giá trị nhỏ nhất thỏa mãn điều kiện cho trước")


*Như vậy, bài viết trên đã hướng dẫn các bạn sử dụng hàm DMIN để tìm giá trị nhỏ nhất thỏa mãn điều kiện cho trước trong Excel. Chúc các bạn thành công!*


moreNếu Hàm MIN dùng để tìm giá trị nhỏ nhất của các giá trị thì…


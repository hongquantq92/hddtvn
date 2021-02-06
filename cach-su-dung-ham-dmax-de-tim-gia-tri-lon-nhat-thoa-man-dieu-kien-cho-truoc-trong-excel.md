Cách sử dụng hàm DMAX để tìm giá trị lớn nhất thỏa mãn điều kiện cho trước trong Excel
======================================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-dmax-de-tim-gia-tri-lon-nhat-thoa-man-dieu-kien-cho-truoc-trong-excel/)Hàm DMAX trong Excel dùng để tìm kiếm giá trị lớn nhất có điều kiện kèm theo trong một cột hay một hàng của một dánh sách, hay một bảng dữ liệu nào đó. Người dùng sẽ tìm giá trị lớn nhất trong một trường dữ liệu Field thuộc bảng tính dữ liệu Database, nhằm …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm DMAX trong Excel dùng để tìm kiếm giá trị lớn nhất có điều kiện kèm theo trong một cột hay một hàng của một dánh sách, hay một bảng dữ liệu nào đó. Người dùng sẽ tìm giá trị lớn nhất trong một trường dữ liệu Field thuộc bảng tính dữ liệu Database, nhằm thỏa mãn một điều kiện cho trước Criterial mà người dùng chọn. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm DMAX trong Excel.**


![](https://hddtvn.com/wp-content/uploads/2021/01/DMAX.png)


### 1. Cấu trúc hàm DMAX


Cú pháp hàm: **=DMAX(database; field; criteria)**


*Trong đó:*




* **database**: đối số bắt buộc, danh sách hay cơ sở dữ liệu có liên quan bao gồm cả các Tiêu đề cột.

* **field**: đối số bắt buộc, trường (cột) cần lấy giá trị lớn nhất. Các bạn có thể nhập trực tiếp tiêu đề cột trong dấu ngoặc kép hay một số thể hiện vị trí cột trong **database** hoặc cũng có thể nhập ô chứa tiêu đề cột cần sử dụng.

* **criteria**: đối số bắt buộc, là phạm vi ô chứa điều kiện. Các bạn có thể chọn bất kỳ miễn sao phạm vi đó chứa ít nhất một Tiêu đề cột và ô dưới tiêu đề cột chứa điều kiện cho cột.



*Lưu ý:*




* Nên đặt phạm vi điều kiện criteria trên trang tính để khi thêm dữ liệu thì phạm vi chứa điều kiện không thay đổi.

* Mặc dù phạm vi điều kiện có thể được định vị ở bất kỳ đâu trên trang tính. Nhưng phạm vi điều kiện cần tách rời không chèn lên danh sách hay cơ sở dữ liệu cần xử lý.

* **criteria** bắt buộc phải chứa ít nhất Tiêu đề cột và một ô chứa điều kiện dưới tiêu đề cột.



### 2. Cách sử dụng hàm DMAX


Ví dụ ta có bảng dữ liệu như bảng dưới và yêu cầu tìm tiền hàng lớn nhất của loại hàng là Đá


![](https://hddtvn.com/wp-content/uploads/2021/01/OMEg0eD.png)


Đầu tiên ta cần tạo điều kiện loại hàng là Đá


![](https://hddtvn.com/wp-content/uploads/2021/01/af6iSH4.png)


Tiếp theo ta áp dụng công thức DMAX:


**=DMAX(A1:F11;F1;D14:D15)**


*Trong đó:*




* **A1:F11** là bảng dữ liệu

* **F1** là cột TIỀN HÀNG, cột cần tìm giá trị cao nhất.

* **D14:D15** là vùng điều kiện criteria.



Kết quả ta sẽ thu được tiền hàng cao nhất của Đá như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/SBkvtRL.png)


Hoặc cũng với công thức trên, ta có thể thay **field** thành số 6 là vị trí của cột TIỀN HÀNG trong bảng.


**=DMAX(A1:F11;6;D14:D15)**


Ta cũng thu được kết quả tương tự như trên


![](https://hddtvn.com/wp-content/uploads/2021/01/IbHIjRF.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn sử dụng hàm DMAX để tìm giá trị lớn nhất thỏa mãn điều kiện cho trước trong Excel. Chúc các bạn thành công!*


moreHàm DMAX trong Excel dùng để tìm kiếm giá trị lớn nhất có điều kiện…


Cách sử dụng hàm DSUM để tính tổng có điều kiện trong Excel
===========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-dsum-de-tinh-tong-co-dieu-kien-trong-excel/)Hàm DSUM là hàm tính tổng của một trường hoặc 1 cột để thỏa mãn điều kiện đưa ra. Công thức và cách sử dụng hàm DSUM cũng đơn giản và nhanh hơn so với hàm SUMIF. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm DSUM trong Excel. 1. Cấu …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm DSUM là hàm tính tổng của một trường hoặc 1 cột để thỏa mãn điều kiện đưa ra. Công thức và cách sử dụng hàm DSUM cũng đơn giản và nhanh hơn so với hàm SUMIF. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm DSUM trong Excel.**


### 1. Cấu trúc hàm DSUM


Cú pháp hàm: **=DSUM(database;field;criteria)**


*Trong đó:*




* **Database** là cơ sở dữ liệu được tạo từ 1 phạm vi ô. Danh sách dữ liệu này sẽ chứa các dữ liệu là các trường, gồm trường để kiểm tra điều kiện và trường để tính tổng. Danh sách chứa hàng đầu tiên là tiêu đề cột.

* **Field** chỉ rõ tên cột dùng để tính tổng các số liệu. Có thể nhập tên tiêu đề cột trong dấu ngoặc kép hoặc là một số thể hiện vị trí cột trong danh sách không trong dấu ngoặc kép (ví dụ số 1 là cột đầu tiên, số 2 là cột thứ 2… trong database) hoặc tham chiếu đến tiêu đề cột mà các bạn muốn tính tổng.

* **Criteria** là phạm vi ô có chứa điều kiện mà các bạn muốn hàm DSUM kiểm tra.



*Lưu ý về phạm vi ô chứa điều kiện **Criteria***




* Các bạn có thể dùng phạm vi bất kỳ nào cho đối số **Criteria** nếu phạm vi đó có chứa ít nhất một nhãn cột và ít nhất một ô bên dưới tiêu đề cột đó mà nó sẽ xác định điều kiện cho cột đó.

* Không nên đặt phạm vi điều kiện ở phía dưới danh sách, vì sẽ không có vị trí thêm các thông tin khác vào danh sách.



### 2. Cách sử dụng hàm DSUM


Ví dụ ta có bảng dữ liệu như sau


![](https://hddtvn.com/wp-content/uploads/2021/01/UiJU4DD.png)


#### a. Tính tổng doanh thu của mặt hàng bia trong năm 2017


Đầu tiên ta có thể tạo một phạm vi điều kiện cho hàm DSUM. Vì cần tính mặt hàng bia mà có rất nhiều loại bia nên điều kiện ta sẽ nhập thêm ký tự đại diện là dấu * sau Bia.


![](https://hddtvn.com/wp-content/uploads/2021/01/F8CzYQG.png)


Tiếp theo các bạn nhập công thức hàm DSUM như sau:


**=DSUM(A1:E10;C1;G4:G5)**


*Trong đó:*




* **A1:E10** là phạm vi cơ sở dữ liệu tại đây chứa cột cần tính tổng và cột chứa điều kiện cần kiểm tra.

* **C1** là tiêu đề cột mà ta sẽ sử dụng giá trị trong cột đó để tính tổng.

* **G4:G5** là phạm vi điều kiện chứa tiêu đề cột và một giá trị điều kiện.



![](https://hddtvn.com/wp-content/uploads/2021/01/YGObF6s.png)


Hoặc ta cũng có thể thay giá trị trong Field thành tên của ô tiêu đề của cột mà ta sẽ sử dụng giá trị trong cột đó để tính tổng như sau:


**=DSUM(A1:E10;”Doanh thu 2017″;G4:G5)**


Ta thu được kết quả vẫn như trên


![](https://hddtvn.com/wp-content/uploads/2021/01/0X8VpKc.png)


#### b. Tính doanh thu năm 2018 của những mặt hàng có doanh thu năm 2018 lớn hơn 500.000.000.


Đầu tiên ta cũng tạo phạm vi điều kiện với tiêu đề cột Doanh thu 2018, và giá trị điều kiện là >500000000.


![](https://hddtvn.com/wp-content/uploads/2021/01/Te9KACw.png)


Ta nhập công thức hàm DSUM:


**=DSUM(A1:E10;D1;G4:G5)**


*Trong đó:*




* **A1:E10** là phạm vi cơ sở dữ liệu mà các bạn sẽ làm việc.

* **D1** là tham chiếu đến tiêu đề cột cần tính tổng.

* **G4:G5** là phạm vi điều kiện criteria với tiêu đề cột Doanh thu 2018 và điều kiện là >500000000.



![](https://hddtvn.com/wp-content/uploads/2021/01/IMo1zDe.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DSUM trong Excel. Chúc các bạn thành công!*


moreHàm DSUM là hàm tính tổng của một trường hoặc 1 cột để thỏa mãn…


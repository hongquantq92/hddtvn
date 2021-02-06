Cách dùng hàm FIXED để làm tròn đến một vị trí số thập phân trong Excel
=======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-fixed-de-lam-tron-den-mot-vi-tri-so-thap-phan-trong-excel/)Hàm FIXED trong excel trả về một đại diện văn bản của một số được làm tròn đến một số vị trí thập phân được chỉ định. FIXED là một hàm dựng sẵn trong excel được phân loại là Hàm Chuỗi/Văn bản. Nó có thể được sử dụng như một hàm bảng tính (WS) trong …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm FIXED trong excel trả về một đại diện văn bản của một số được làm tròn đến một số vị trí thập phân được chỉ định. FIXED là một hàm dựng sẵn trong excel được phân loại là Hàm Chuỗi/Văn bản. Nó có thể được sử dụng như một hàm bảng tính (WS) trong Excel. Là một hàm trang tính, hàm FIXED có thể được nhập như một phần của công thức trong một ô của trang tính. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm FIXED trong Excel.**


### 1. Cấu trúc hàm FIXED


Cú pháp hàm: **=FIXED(number; [decimals]; [no\_commas])**


*Trong đó:*




* **number**: Đối số bắt buộc, là số mà bạn muốn làm tròn và chuyển thành văn bản.

* **decimals**: Đối số tùy chọn, là số chữ số nằm bên phải dấu thập phân.

* **no\_commas**: Đối số tùy chọn, là giá trị logic được trả về nếu chuỗi văn bản trả về bao gồm dấu phẩy.



*Lưu ý:*




* Nếu **decimals** bị bỏ qua, nó sẽ lấy giá trị mặc định là 2.

* Nếu **decimals** là số âm, số thập phân sẽ được làm trong từ số bên trái dấu thập phân.

* **no\_commas** là TRUE nếu dấu phẩy không được bao gồm trong chuỗi văn bản trả về.

* **no\_commas** là FALSE nếu dấu phẩy được bao gồm trong chuỗi văn bản trả về.

* Nếu đối số **no\_commas** bị bỏ qua, nó sẽ sử dụng giá trị mặc định là FALSE.

* Số trong Microsoft Excel không bao giờ được có quá 15 chữ số có nghĩa, nhưng đối số thập phân có thể lớn đến 127.

* Sự khác nhau chính giữa việc định dạng một ô có chứa số bằng lệnh và định dạng số trực tiếp bằng hàm FIXED là hàm FIXED chuyển đổi kết quả của nó thành văn bản. Số được định dạng bằng lệnh Ô vẫn là số.

* Hàm FIXED sẽ trả về kết quả lỗi #VALUE! nếu đối số được cung cấp không phải kiểu số.



### 2. Cách sử dụng hàm FIXED


Ví dụ ta cần chuyển đổi số 345,678 thành văn bản và làm tròn đến các số vị trí thập phân khác nhau như sau:




* Làm tròn số trong ô A2 sang trái dấu thập phân hai chữ số

* Làm tròn số trong ô A2 sang phải dấu thập phân một chữ số

* Làm tròn số trong ô A2 sang trái dấu thập phân một chữ số

* Làm tròn số trong ô A3 sang trái dấu thập phân một chữ số, không có dấu phẩy

* Làm tròn số trong ô A3 sang phải dấu thập phân một chữ số, có bao gồm dấu phẩy



![](https://hddtvn.com/wp-content/uploads/2021/01/Xb9WZ7Q.png)


Ta có công thức cho từng trường hợp như sau:




* Làm tròn số trong ô A2 sang trái dấu thập phân hai chữ số: **=FIXED(A1)**

* Làm tròn số trong ô A2 sang phải dấu thập phân một chữ số: **=FIXED(A1;1)**

* Làm tròn số trong ô A2 sang trái dấu thập phân một chữ số: **=FIXED(A1;-2)**

* Làm tròn số trong ô A3 sang trái dấu thập phân một chữ số, không có dấu phẩy: **=FIXED(A1;-1;TRUE)**

* Làm tròn số trong ô A3 sang phải dấu thập phân một chữ số, có bao gồm dấu phẩy: **=FIXED(A1;1;FALSE)**



Ta thu được kết quả như hình dưới:


![](https://hddtvn.com/wp-content/uploads/2021/01/7sEKoov.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm FIXED để làm tròn số trong Excel. Chúc các bạn thành công!*


moreHàm FIXED trong excel trả về một đại diện văn bản của một số được…


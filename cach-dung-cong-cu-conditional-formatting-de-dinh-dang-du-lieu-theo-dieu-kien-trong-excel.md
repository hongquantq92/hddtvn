Cách dùng công cụ Conditional Formatting để định dạng dữ liệu theo điều kiện trong Excel
========================================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-cong-cu-conditional-formatting-de-dinh-dang-du-lieu-theo-dieu-kien-trong-excel/)Định dạng có điều kiện (Conditional Formatting) là một công cụ mạnh mẽ và có thể thay đổi “diện mạo” của một ô dựa trên giá trị của nó, giúp người xem xác định dữ liệu quan trọng nhanh chóng. Bạn có thể sử dụng Conditional Formatting để định dạng dữ liệu theo nhiều cách …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Định dạng có điều kiện (Conditional Formatting) là một công cụ mạnh mẽ và có thể thay đổi “diện mạo” của một ô dựa trên giá trị của nó, giúp người xem xác định dữ liệu quan trọng nhanh chóng. Bạn có thể sử dụng Conditional Formatting để định dạng dữ liệu theo nhiều cách khác nhau: thay đổi màu sắc, màu chữ, các kiểu đường viền của một hoặc nhiều ô, hàng, cột hoặc toàn bộ bảng. Định dạng của một ô/cột/hàng sẽ thay đổi dựa theo các quy tắc (điều kiện) mà bạn đã xác định trước. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng công cụ Conditional Formatting để định dạng dữ liệu theo điều kiện, nhấn mạnh hoặc phân biệt giữa dữ liệu và thông tin được lưu trữ trong Excel.**


### 1. Tô màu ô thỏa mãn điều kiện (Highlight Cells Rules)


Ví dụ ta có bảng dữ liệu doanh thu như sau:


[![](https://hddtvn.com/wp-content/uploads/2021/01/0lgIBpX.png)](https://hddtvn.com/wp-content/uploads/2021/01/0lgIBpX.png)


Đầu tiên ta cần bôi đen vùng dữ liệu cần định dạng. Sau đó chọn thẻ **Home** => chọn biểu tượng **Conditional Formatting** => chọn **Highlight Cells Rules**.


![](https://hddtvn.com/wp-content/uploads/2021/01/h2NqN7K.png)


Tại đây ta có thể thấy các mục điều kiện có sẵn là:




* **Greater than**: Giá trị lớn hơn.

* **Less than**: Giá trị nhỏ hơn.

* **Between**: Giá trị trong khoảng.

* **Equal to**: Giá trị bằng.

* **Text that Contains**: Giá trị chứa một chuỗi ký tự nào đó.

* **A Date Occciring**: Giá trị chứa ngày tháng định sẵn.

* **Duplicate Value**: Giá trị trùng lặp.



Ví dụ nếu cần tìm các giá trị doanh thu trong khoảng 500.000.000 đến 1.000.000.000 thì ta chọn mục Between


![](https://hddtvn.com/wp-content/uploads/2021/01/zLNdYAT.png)


Sau khi cửa sổ Between hiện ra, tại mục giá trị so sánh bạn gõ số 500.000.000 và 1.000.000.000. Ta có thể lựa chọn các định dạng có sẵn tại mục định dạng bôi màu hoặc tùy chỉnh bằng cách chọn Custom Format.


![](https://hddtvn.com/wp-content/uploads/2021/01/5y9yyWw.png)


Sau khi lựa chọn xong các tùy chọn bạn nhấn nút OK. Và kết quả là các ô chứa giá trị từ 500.000.000 đến 1.000.000.000 được bôi màu.


![](https://hddtvn.com/wp-content/uploads/2021/01/C9zjUEa.png)


### 2. Tô màu ô theo thứ hạng (Top/Bottom Rules)


Tại mục này có các điều kiện có sẵn như sau:




* **Top 10 Items**: Định dạng 10 ô có giá trị lớn nhất.

* **Top 10%**: Định dạng 10% số lượng ô có giá trị lớn nhất.

* **Bottom 10 Items**: Định dạng 10 ô có giá nhị nhỏ nhất.

* **Bottom 10%**: Định dạng 10% số lượng ô có giá trị nhỏ nhất.

* **Above Average**: Định dạng ô lớn hơn giá trị trung bình của vùng dữ liệu.

* **Below Average**: Định dạng ô nhỏ hơn giá trị trung bình của vùng dữ liệu.



![](https://hddtvn.com/wp-content/uploads/2021/01/go7m958.png)


Ví dụ nếu ta cần định dạng 10% ô có giá trị lớn nhất, thì ta chọn mục **Top 10%**.


![](https://hddtvn.com/wp-content/uploads/2021/01/HHtTRC8.png)


### 3. Định dạng mức độ giá trị của từng ô (Data Bar)


Với định dạng này, mức độ lớn nhỏ của từng ô trong vùng dữ liệu sẽ được phân định bằng việc bôi màu dài hay ngắn.


![](https://hddtvn.com/wp-content/uploads/2021/01/Q2WkeOG.png)


### 4. Sắp xếp dữ liệu tăng giảm theo độ đậm nhạt của màu sắc (Color Scale)


Mục này sẽ sắp xếp thứ tự của từng ô trong chuỗi và bôi màu theo định dạng mà người dùng lựa chọn.


![](https://hddtvn.com/wp-content/uploads/2021/01/mQsn63H.png)


### 5. Thêm biểu tượng vào ô giá trị (Icon Sets)


Với kiểu mẫu định dạng này, Excel phân nhóm dữ liệu bằng biểu tượng đặc biệt như mũi tên, đèn giao thông,..


![](https://hddtvn.com/wp-content/uploads/2021/01/U1txOvV.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách dùng công cụ Conditional Formatting để định dạng dữ liệu theo điều kiện trong Excel. Chúc các bạn thành công!*


moreĐịnh dạng có điều kiện (Conditional Formatting) là một công cụ mạnh mẽ và có…


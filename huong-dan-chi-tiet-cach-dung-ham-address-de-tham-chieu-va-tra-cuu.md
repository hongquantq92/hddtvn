Hướng dẫn chi tiết cách dùng hàm ADDRESS để tham chiếu và tra cứu
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-chi-tiet-cach-dung-ham-address-de-tham-chieu-va-tra-cuu/)Hàm ADDRESS là một trong những hàm tham chiếu và tra cứu thường xuyên được sử dụng trong Excel. Trong bài viết này, hddtvn sẽ giới thiệu đến bạn đọc cách sử dụng hàm ADDRESS. 1. Cấu trúc hàm Cú pháp hàm: =ADDRESS(row\_num; column\_num; [abs\_num]; [a1]; [sheet\_text]) Trong đó: row\_num: đối số bắt buộc, là …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm ADDRESS là một trong những hàm tham chiếu và tra cứu thường xuyên được sử dụng trong Excel. Trong bài viết này, hddtvn sẽ giới thiệu đến bạn đọc cách sử dụng hàm ADDRESS.**


### 1. Cấu trúc hàm


Cú pháp hàm: **=ADDRESS(row\_num; column\_num; [abs\_num]; [a1]; [sheet\_text])**


*Trong đó:*




* **row\_num**: đối số bắt buộc, là giá trị số chỉ rõ số hàng dùng trong tham chiếu ô.

* **column\_num**: đối số bắt buộc, là giá trị số chỉ rõ số cột dùng trong tham chiếu ô.

* **abs\_num**: đối số tùy chọn, là giá trị số chỉ rõ kiểu tham chiếu cần trả về.

* **A1**: đối số tùy chọn, là giá trị logic chỉ rõ kiểu tham chiếu A1 hay R1C1. Trong kiểu A1, các cột được đánh nhãn theo bảng chữ cái, còn các hàng được đánh nhãn dạng số. Trong kiểu R1C1, cả cột và hàng đều được đánh nhãn dạng số.

* **sheet\_text**: đối số tùy chọn, là giá trị văn bản chỉ rõ tên của trang tính được dùng làm tham chiếu ngoài.



*Lưu ý:*




* Nếu đối số **abs\_num** là 1 hoặc bỏ qua thì hàm trả về kiểu tham chiếu tuyệt đối. Nếu **abs\_num** là 2 thì hàm trả về kiểu tham chiếu hàng tuyệt đối, cột tương đối. Nếu **abs\_num** là 3 thì hàm trả về kiểu tham chiếu hàng tương đối, cột tuyệt đối. Nếu đối số **abs\_num** là 4 thì hàm trả về kiểu tham chiếu tương đối.

* Nếu đối số **A1** là **TRUE** hoặc được bỏ qua, thì hàm **ADDRESS** trả về tham chiếu kiểu **A1**; nếu **FALSE**, thì hàm **ADDRESS** trả về tham chiếu kiểu **R1C1**.

* Nếu đối số **sheet\_text** được bỏ qua, thì không có tên trang tính nào được dùng và địa chỉ mà hàm trả về tham chiếu tới một ô trên trang tính hiện tại.



### 2. Cách sử dụng hàm ADDRESS


Áp dụng cấu trúc hàm bên trên, ta có ví dụ những công thức cơ bản như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/q4fwcUj.png)


*Lưu ý:*




* Kiểu tham chiếu A1 ( dòng 1 và cột A): tham chiếu tuyệt đối được biểu hiện bằng kí tự $ đứng trước vị trí cột hoặc dòng.

* Kiểu tham chiếu R1C1 (dòng 1 và cột 1): tham chiếu tuyệt đối được biểu hiện bằng dấu ngoặc vuông [] ở số dòng hoặc số cột.

* Nếu bạn muốn tham chiếu đến một cửa sổ làm việc (Workbook), và trang tính khác bạn để tên của cửa sổ làm việc trong dấu ngoặc vuông. Ví dụ “[Book1]Sheet2” có nghĩa là cửa sổ làm việc có tên là Book1 và trang tính trong cửa sổ đó là Sheet2.



*Như vậy, bài viết trên đã giới thiệu đến các bạn cách sử dụng hàm ADDRESS trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm ADDRESS là một trong những hàm tham chiếu và tra cứu thường xuyên được…


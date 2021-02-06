Cách sử dụng hàm OR để kiểm tra nhiều điều kiện cùng một lúc
============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-or-de-kiem-tra-nhieu-dieu-kien-cung-mot-luc/)Hàm OR là một hàm logic để kiểm tra nhiều điều kiện cùng một lúc. OR trả về kết quả TRUE hoặc FALSE. Giả sử bạn muốn kiểm tra ô A2 có đạt điều kiện lớn hơn 1 và nhỏ hơn 1000. Nếu A2 có giá trị lớn hơn 1 HOẶC nhỏ hơn 100 thì …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm OR là một hàm logic để kiểm tra nhiều điều kiện cùng một lúc. OR trả về kết quả TRUE hoặc FALSE. Giả sử bạn muốn kiểm tra ô A2 có đạt điều kiện lớn hơn 1 và nhỏ hơn 1000. Nếu A2 có giá trị lớn hơn 1 HOẶC nhỏ hơn 100 thì kết quả trả về là TRUE, nếu không kết quả trả về là FALSE. Trong bài viết này, hddtvn sẽ giới thiệu tới bạn đọc hàm OR và cách kết hợp với các hàm logic khác trong Excel.**


### 1. Cấu trúc hàm


Cú pháp hàm: **=OR(logical1; [logical2]; …)**


*Trong đó:*




* **logical1**: đối số bắt buộc, là mệnh đề logic đầu tiên.

* **logical2, …**: đối số tùy chọn, là các mệnh đề logic tiếp theo.



*Lưu ý:*




* Hàm OR cho phép sử dụng nhiều hơn 1 mệnh đề logic và tối đa là 256 mệnh đề

* Kết quả của hàm OR là TRUE nếu một hoặc tất cả các mệnh đề logic đều đúng

* Kết quả của hàm OR là FALSE nếu tất cả mệnh đề logic đều sai



### 2. Cách sử dụng hàm OR


Ví dụ ta có bảng điểm thi môn toán và văn của học sinh như hình dưới. Nếu các học sinh có cả điểm thi toán và văn đều nhỏ hơn 5 thì sẽ trượt.


[![](https://hddtvn.com/wp-content/uploads/2021/01/WYMupWM.png)](https://hddtvn.com/wp-content/uploads/2021/01/WYMupWM.png)


Áp dụng cấu trúc hàm như trên, ta có công thức để tìm các học sinh có điểm thi toán và văn đều nhỏ hơn 5 như sau:


**=OR(D2>5;E2>5)**


Sao chép công thức cho tất cả học sinh ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/fAeQYrF.png)


Như kết quả ta có thể thấy có 3 bạn hàm trả về kết quả FALSE tức là cả điểm thi toán và văn đều nhỏ hơn 5 là Tú, Dũng và Hùng.


Ta cũng có thể sử dụng hàm OR kết hợp với các hàm logic khác như AND và IF để tính toán. Ví dụ nếu bạn nào giới tính là Nam và có một trong hai điểm toán hoặc văn nhỏ hơn 5 thì Trượt, còn lại Đỗ. Ta có công thức tìm như sau:


**=IF(AND(C2=”Nam”;OR(D2<5;E2<5));”Trượt”;”Đỗ”)**


Sao chép công thức cho tất cả học sinh ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/SsDHg42.png)


Nhìn vào kết quả ta có thể thấy, có 4 bạn trượt là Dũng, Thành, Hùng và Hoàng do có giới tính là Nam và có 1 trong 2 môn toán hoặc văn có điểm thấp hơn 5.


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm OR và cách kết hợp hàm OR với hàm IF và AND trong Excel. Chúc các bạn thành công!*


moreHàm OR là một hàm logic để kiểm tra nhiều điều kiện cùng một lúc….


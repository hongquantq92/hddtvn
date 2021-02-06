Mẹo dùng hàm AND và kết hợp với các hàm logic khác trong Excel
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-dung-ham-and-va-ket-hop-voi-cac-ham-logic-khac-trong-excel/)Hàm AND là một hàm logic cơ bản được sử dụng nhiều trong Excel. Hàm này thường được sử dụng kèm theo với hàm IF để xét nhiều logic cùng lúc. Hàm sẽ trả về kết quả là TRUE nếu tất cả các đối số của hàm định trị là TRUE và trả về kết …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm AND là một hàm logic cơ bản được sử dụng nhiều trong Excel. Hàm này thường được sử dụng kèm theo với hàm IF để xét nhiều logic cùng lúc. Hàm sẽ trả về kết quả là TRUE nếu tất cả các đối số của hàm định trị là TRUE và trả về kết quả là FALSE nếu có ít nhất một đối số của hàm định trị là FALSE. Trong bài viết này, hddtvn sẽ giới thiệu tới bạn đọc hàm AND và cách kết hợp hàm AND với các hàm logic khác trong Excel.**


### 1. Cấu trúc hàm AND


Cú pháp hàm: **=AND(logical1; [logical2]; …)**


*Trong đó:*




* **logical1**: đối số bắt buộc, là mệnh đề logic đầu tiên.

* **logical2, …**: đối số tùy chọn, là các mệnh đề logic tiếp theo.



*Lưu ý:*




* Hàm AND cho phép sử dụng nhiều hơn 1 mệnh đề logic và tối đa là 256 mệnh đề

* Kết quả của hàm AND là True nếu tất cả các mệnh đề logic đều đúng

* Kết quả của hàm AND là False nếu một mệnh đề logic bất kỳ nào trong hàm sai



### 2. Cách sử dụng hàm AND


Ví dụ ta có bảng điểm thi môn toán và văn của học sinh như hình dưới. Yêu cầu tìm các học sinh có điểm thi toán và văn đều lớn hơn 5.


![](https://hddtvn.com/wp-content/uploads/2021/01/53BcOlU.png)


Áp dụng cấu trúc hàm như trên, ta có công thức để tìm các học sinh có điểm thi toán và văn đều lớn hơn 5 như sau:


**=AND(D2>5;E2>5)**


Sao chép công thức cho tất cả học sinh ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/EfAp0bi.png)


Như kết quả ta có thể thấy chỉ có 4 bạn hàm trả về kết quả TRUE tức là cả điểm thi toán và văn đều lớn hơn 5 là Huy, Huyền, Hương và Hùng.


Ta cũng có thể sử dụng hàm AND kết hợp với các hàm logic khác để tính toán. Ví dụ trong trường hợp này là kết hợp hàm AND với hàm IF để trả về kết quả là bạn có điểm văn và toán > 5 thì “Đỗ”. Còn bạn nếu có điểm văn hoặc toán <5 thì “Trượt”.


**=IF(AND(D2>5;E2>5);”Đỗ”;”Trượt”)**


Sao chép công thức cho tất cả học sinh ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/uhN27Rq.png)


Như kết quả thu được, ta có thể thấy 4 bạn có kết quả Đỗ là Huy, Huyền, Hương và Hùng với cả điểm văn và toán đều lớn hơn 5.


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm AND và cách kết hợp hàm AND với hàm IF trong Excel. Chúc các bạn thành công!*


moreHàm AND là một hàm logic cơ bản được sử dụng nhiều trong Excel. Hàm…


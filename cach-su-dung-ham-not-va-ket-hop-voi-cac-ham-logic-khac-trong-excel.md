Cách sử dụng hàm NOT và kết hợp với các hàm logic khác trong Excel
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-not-va-ket-hop-voi-cac-ham-logic-khac-trong-excel/)Hàm NOT là một trong những hàm logic cơ bản được sử dụng rất phổ biến trong Excel. Đây là hàm trả về giá trị nghịch đảo của một giá trị mà bạn đã biết kết quả của nó. Ví dụ bạn có một bảng dữ liệu lớn và muốn tìm các giá trị lớn …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm NOT là một trong những hàm logic cơ bản được sử dụng rất phổ biến trong Excel. Đây là hàm trả về giá trị nghịch đảo của một giá trị mà bạn đã biết kết quả của nó.** **Ví dụ bạn có một bảng dữ liệu lớn và muốn tìm các giá trị lớn 500. Hàm NOT giúp bạn tìm ra dữ liệu nào có kết quả Đạt hay Không đạt.** **Trong bài viết sau đây, hddtvn sẽ giới thiệu tới bạn chi tiết cách dùng hàm NOT và cách kết hợp với các hàm logic khác.**


### 1. Cấu trúc hàm


Cú pháp hàm: **=NOT(logical)**


*Trong đó:* **logical** là một giá trị hoặc mệnh đề logic.


*Lưu ý:*




* Hàm NOT trả về giá trị nghịch đảo của một giá trị hoặc mệnh đề logic.

* Nếu mệnh đề logic là TRUE, hàm NOT sẽ trả về FALSE.

* Nếu mệnh đề logic là FALSE, hàm NOT sẽ trả về TRUE.



### 2. Cách sử dụng hàm NOT


Ví dụ ta có bảng điểm thi môn toán và văn của học sinh như hình dưới.


[![](https://hddtvn.com/wp-content/uploads/2021/01/WYMupWM.png)](https://hddtvn.com/wp-content/uploads/2021/01/WYMupWM.png)


Để tìm các bạn học sinh có điểm thi toán lớn hơn hoặc bằng 5 ta có công thức tại ô F2 như sau:


**=NOT(D2<5)**


Sao chép công thức cho tất cả học sinh ta thu được kết quả các bạn có điểm thi toán lớn hơn 5 là TRUE và các bạn có điểm thi toán nhỏ hơn 5 là FALSE


![](https://hddtvn.com/wp-content/uploads/2021/01/mNe6mr3.png)


Ta cũng có thể sử dụng hàm NOT kết hợp với các hàm logic khác như AND và IF để tính toán. Ví dụ nếu bạn nào có cả điểm toán và văn nhỏ hơn 5 thì Trượt, còn lại Đỗ. Ta có công thức tìm như sau:


**=IF(AND(NOT(D2>5);NOT(E2>5));”Trượt”;”Đỗ”)**


Sao chép công thức cho tất cả học sinh ta thu được kết quả là 3 bạn Tú, Dũng và Hùng trượt do có cả điểm văn và toán đều nhỏ hơn 5.


![](https://hddtvn.com/wp-content/uploads/2021/01/PCWlwtS.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm NOT và cách kết hợp hàm NOT với hàm IF và AND trong Excel. Chúc các bạn thành công!*


moreHàm NOT là một trong những hàm logic cơ bản được sử dụng rất phổ…


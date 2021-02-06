Cách dùng hàm CONFIDENCE để tính khoảng tin cậy của trung bình tổng thể
========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-confidence-de-tinh-khoa%cc%89ng-tin-cay-cua-trung-binh-tong-the/)Hàm CONFIDENCE là một hàm toán học và lượng giác được sử dụng trong Excel để tính khoảng tin cậy của trung bình tổng thể. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm CONFIDENCE trong Excel. 1. Cấu trúc hàm CONFIDENCE Cú pháp hàm: =CONFIDENCE(alpha; standard\_dev; size) Trong đó: Alpha: …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm CONFIDENCE là một hàm toán học và lượng giác được sử dụng trong Excel để tính khoảng tin cậy của trung bình tổng thể. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm CONFIDENCE trong Excel.**


### 1. Cấu trúc hàm CONFIDENCE


Cú pháp hàm: **=CONFIDENCE(alpha; standard\_dev; size)**


*Trong đó:*




* **Alpha**: đối số bắt buộc, là mức quan trọng được dùng để tính toán mức tin cậy. Mức tin cậy = 100%*(1-  alpha). Tức là nếy alpha = 0,05 thì mức tin cậy = 95%.

* **Standard\_dev**: đối số bắt buộc, là độ lệch chuẩn tổng thể cho phạm vi dữ liệu và được giả định là đã được xác định.

* **Size**: đối số bắt buộc, là cỡ mẫu.



*Lưu ý:*




* Khoảng tin cậy là một phạm vi các giá trị mà trung độ mẫu của bạn nằm ở trung tâm của phạm vi này.

* Hàm CONFIDENCE sẽ trả về giá trị lỗi #VALUE! nếu có bất kỳ đối số nào không phải dạng số.

* Nếu **alpha** ≤ 0 hoặc ≥ 1 thì hàm CONFIDENCE sẽ trả về giá trị lỗi #NUM! .

* Nếu **standard\_dev** ≤ 0 thì CONFIDENCE sẽ trả về giá trị lỗi #NUM! .

* Nếu **size** không phải là số nguyên thì nó bị cắt phần phân số thành số nguyên

* Nếu **size** < 1 thì hàm CONFIDENCE trả về giá trị lỗi #NUM! .



### 2. Cách sử dụng hàm CONFIDENCE


Ví dụ ta có bảng mẫu các tham số của hàm CONFIDENCE như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/sydMlNA.png)


Áp dụng cấu trúc hàm bên trên, ta có công thức tính khoảng tin cậy của trung bình tổng thể như sau:


**=CONFIDENCE(B3;B4;B5)**


*Trong đó:*




* **B3** là mức quan trọng của khoảng tin cậy

* **B4** là độ lệch chuẩn của tổng thể

* **B5** là cỡ mẫu



![](https://hddtvn.com/wp-content/uploads/2021/01/efXsbgc.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm CONFIDENCE trong Excel. Chúc các bạn thành công!*


moreHàm CONFIDENCE là một hàm toán học và lượng giác được sử dụng trong Excel…


Cách dùng hàm IF và cách kết hợp với các hàm khác trong Excel
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-if-va-cach-ket-hop-voi-cac-ham-khac-trong-excel/)Hàm IF thường được dùng để kiểm tra một điều kiện và trả về một giá trị nếu điều kiện được đáp ứng. Và sẽ trả về một giá trị khác nếu điều kiện đó không được đáp ứng. Hàm IF được sử dụng rất phổ biến trong Excel. Nó có thể được sử dụng …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm IF thường được dùng để kiểm tra một điều kiện và trả về một giá trị nếu điều kiện được đáp ứng. Và sẽ trả về một giá trị khác nếu điều kiện đó không được đáp ứng. Hàm IF được sử dụng rất phổ biến trong Excel. Nó có thể được sử dụng đơn lẻ hoặc kết hợp với các hàm khác để phục vụ yêu cầu của người dùng. Bài viết sau đây sẽ hước dẫn các bạn cách dùng hàm IF sao cho hiệu quả nhất.**


### 1. Chú pháp hàm IF


**IF (logical\_test, [value\_if\_true], [value\_if\_false])**


Trong đó:


logical\_test: Là một giá trị hay biểu thức logic có giá trị TRUE (đúng) hoặc FALSE (sai). Bắt buộc phải có. Đối với tham số này, bạn có thể chỉ rõ đó là ký tự, ngày tháng, con số hay bất cứ biểu thức so sánh nào.


Value\_if\_true: Là giá trị mà hàm sẽ trả về nếu biểu thức logic cho giá trị TRUE hay nói cách khác là điều kiện thỏa mãn. Không bắt buộc phải có.


Value\_if\_false: là giá trị mà hàm sẽ trả về nếu biểu thức logic cho giá trị FALSE hay nói cách khác là điều kiện không thỏa mãn. Không bắt buộc phải có.


### 2. Cách dùng hàm IF


#### a. Sử dụng đơn lẻ


Ví dụ: Chúng ta có các loại đá tiêu thụ trong kỳ. Nếu số lượng tiêu thụ lớn hơn 30.000 thì là tiêu thụ tốt, ít hơn 30.000 thì là tiêu thụ kém.


[![](https://hddtvn.com/wp-content/uploads/2021/01/5CGf6AF.png)](https://hddtvn.com/wp-content/uploads/2021/01/5CGf6AF.png)


Trong trường hợp này chúng ta sẽ sử dụng hàm IF với điều kiện cơ bản nhất là đúng hoặc không đúng. Ở đây chúng ta sẽ gán cho hàm IF điều kiện là số đó phải lớn >=30.000. Điều kiện xảy ra sẽ là 1 trong 2 kết quả.  

Công thức hàm IF: =IF(Điều kiện, “Kết quả nếu đúng”, “kết quả nếu không đúng”)


Như vậy chúng ta sẽ nhập tại ô F2: **=IF(E2>=30000;”Tốt”;”Kém”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/iXICZ26.png)


Kết quả sẽ cho ra là Kém do số lượng ít hơn 30.000. Copy công thức xuống các ô bên dưới chúng ta sẽ được kết quả như hình sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/EXWuXAM.png)


#### b. Sử dụng hàm IF kết hợp với hàm OR và hàm AND


Giả sử chúng ta cần phân loại hàng xuất khẩu và nội địa. Hàng xuất khẩu là Đá chất lượng loại 1 và Bột đá chất lượng loại 2.


![](https://hddtvn.com/wp-content/uploads/2021/01/pmXmUbr.png)


Trong trường hợp này chúng ta sẽ sử dụng hàm **IF** kết hợp với hàm **OR**, **AND** để đạt được điều kiện **Đá loại 1** hoặc **Bột đá loại 2**.


Chúng ta nhập tại ô F2: **=IF(OR(AND(D2=”Đá”;E2=1);AND(D2=”Bột đá”;E2=2));”Xuất khẩu”;”Nội địa”)**


Sau đó copy công thức xuống các ô bên dưới. Kết quả thu được sẽ như hình sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/lYDKFh6.png)


*Như vậy bài viết trên đã hướng dẫn các bạn cách sử dụng hàm IF và cách kết hợp hàm IF với hàm OR và AND. Còn rất nhiều cách kết hợp hàm IF với các hàm khác nữa tùy vào yêu cầu của công việc. Chúc các bạn thành công!*


moreHàm IF thường được dùng để kiểm tra một điều kiện và trả về một…


Hướng dẫn sử dụng hàm ACCRINT để tính tiền lãi cộng dồn
=======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-su-dung-ham-accrint-de-tinh-tien-lai-cong-don/)Excel hỗ trợ rất nhiều hàm khác nhau phục vụ cho tính toán số liệu trong công việc nói chung và ngành tài chính nói riêng. Trong đó có hàm ACCRINT có thể hỗ trợ để tính tiền lãi cộng dồn cho chứng khoán trả lãi định kỳ. Bài viết sau đây sẽ hướng dẫn …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel hỗ trợ rất nhiều hàm khác nhau phục vụ cho tính toán số liệu trong công việc nói chung và ngành tài chính nói riêng. Trong đó có hàm ACCRINT có thể hỗ trợ để tính tiền lãi cộng dồn cho chứng khoán trả lãi định kỳ. Bài viết sau đây sẽ hướng dẫn các bạn sử dụng hàm ACCRINT trong Excel.**


### 1. Cấu trúc hàm ACCRINT


Cú pháp hàm: **=ACCRINT(issue; first\_interest; settlement; rate; par; frequency; [basis]; [calc\_method])**


*Trong đó:*




* **Issue**: đối số bắt buộc, là ngày phát hành chứng khoán.

* **First\_interest**: đối số bắt buộc, là ngày tính lãi đầu tiên của chứng khoán.

* **Settlement**: đối số bắt buộc, là ngày tới hạn của chứng khoán (là ngày sau ngày phát hành chứng khoán khi chứng khoán được giao dịch).

* **Rate**: đối số bắt buộc, là lãi suất hằng năm của chứng khoán.

* **Par**: đối số bắt buộc, là mệnh giá của chứng khoán. Nếu bỏ qua mệnh giá, hàm ACCRINT sẽ dùng $1.000.

* **Frequency**: đối số bắt buộc, là số lần trả lãi hằng năm (Trả n lần mỗi năm thì **frequency** = n)

* **Basis**: đối số tùy chọn, là cơ sở dùng để đếm ngày (mặc định là 0), trong đó:



– Basis = 0 : Một tháng có 30 ngày / Một năm có 360 ngày (theo tiêu chuẩn Bắc Mỹ).  

– Basis = 1 : Số ngày thực tế của mỗi tháng / Số ngày thực tế của mỗi năm.  

– Basis = 2 : Số ngày thực tế của mỗi tháng / Một năm có 360 ngày.  

– Basis = 3 : Số ngày thực tế của mỗi tháng / Một năm có 365 ngày.  

– Basis = 4 : Một tháng có 30 ngày / Một năm có 360 ngày (theo tiêu chuẩn Châu Âu).




* **Calc\_method**: đối số tùy chọn, là giá trị logic chỉ cách để tính số lãi gộp khi **settlement** xảy ra sau **fisrt\_interest**.



– Nếu là 1 (TRUE) thì số lãi gộp sẽ được tính từ ngày phát hành chứng khoán.  

– Nếu là 0 (FALSE) thì số lãi gộp sẽ tính từ ngày tính lãi đầu tiên của chứng khoán.  

– Nếu bỏ qua thì calc\_method mặc định là 1.


#### *Lưu ý:*




* Excel lưu trữ ngày tháng ở dạng số liên tiếp để sử dụng trong tính toán. Theo mặc định, ngày 1 tháng một năm 1900 là số 1 và ngày 1 tháng một năm 2008 là số 39448 bởi nó là 39.448 ngày sau ngày 1 tháng một năm 1900.

* **Issue**, **first\_interest**, **settlement**, **frequency** và cơ sở được cắt cụt thành số nguyên.

* Nếu **issue**, **first\_interest** hoặc **settlement** không phải là ngày hợp lệ, thì hàm ACCRINT trả về giá trị lỗi #VALUE! .

* Nếu lãi suất ≤ 0 hoặc nếu mệnh giá ≤ 0, hàm ACCRINT trả về giá trị lỗi #NUM! .

* Nếu tần suất là một số không phải là 1, 2 hoặc 4, thì hàm ACCRINT trả về giá trị lỗi #NUM! .

* Nếu cơ sở < 0 hoặc nếu cơ sở > 4, thì hàm ACCRINT trả về giá trị lỗi #NUM! .

* Nếu ngày phát hành ≥ ngày thanh toán, thì hàm ACCRINT trả về giá trị lỗi #NUM! .



### 2. Cách sử dụng hàm ACCRINT


Ví dụ ta có bảng tính như hình dưới. Yêu cầu tính tiền lãi cộng dồn cho chứng khoán trả lãi định kỳ với các trường hợp.


![](https://hddtvn.com/wp-content/uploads/2021/01/L9lvape.png)


Với trường hợp giá trị logic của tham số **Calc\_method** là TRUE ta có công thức:


**=ACCRINT(B3;B4;B5;B6;B7;B8;B9;B10)**


![](https://hddtvn.com/wp-content/uploads/2021/01/ePbkobK.png)


Còn với trường hợp Calc\_method là **FALSE** ta có công thức tính:


**=ACCRINT(B3;B4;B5;B6;B7;B8;B9;B11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/l8oevoB.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm ACCRINT để tính tiền lãi cộng dồn cho chứng khoán trả lãi định kỳ. Chúc các bạn thành công!*


moreExcel hỗ trợ rất nhiều hàm khác nhau phục vụ cho tính toán số liệu…


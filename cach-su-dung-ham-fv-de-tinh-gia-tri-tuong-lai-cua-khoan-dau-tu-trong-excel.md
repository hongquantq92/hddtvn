Cách sử dụng hàm FV để tính giá trị tương lai của khoản đầu tư trong Excel
==========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-fv-de-tinh-gia-tri-tuong-lai-cua-khoan-dau-tu-trong-excel/)Hàm FV sẽ tính giá trị tại một thời điểm tương lai nào đó mà các nhà đầu tư sẽ nhận được khi họ đầu tư một kỳ hay nhiều kỳ với lãi suất không đổi. Ví dụ bạn định tiết kiệm một số tiền trong vòng 10 năm và muốn tính xem sau thời …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm FV sẽ tính giá trị tại một thời điểm tương lai nào đó mà các nhà đầu tư sẽ nhận được khi họ đầu tư một kỳ hay nhiều kỳ với lãi suất không đổi. Ví dụ bạn định tiết kiệm một số tiền trong vòng 10 năm và muốn tính xem sau thời gian đó trong tài khoản của bạn sẽ có bao nhiều tiền. Lúc này hãy sử dụng tới hàm FV. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm FV trong Excel.**


### 1. Cấu trúc hàm FV


Cú pháp hàm: **=FV(rate; nper; pmt; [pv]; [type])**


*Trong đó:*




* **Rate**: đối số bắt buộc, là lãi suất theo kỳ hạn (tính theo tháng, quý, năm)

* **Nper**: đối số bắt buộc, là tổng số kỳ hạn thanh toán.

* **Pmt**: đối số bắt buộc, là số tiền thanh toán cho mỗi kỳ. Khoản này là cố định. Thông thường, **pmt** có chứa tiền gốc và lãi, nhưng không chứa các khoản chi phí.

* **Pv**: đối số tùy chọn, là giá trị đầu tư ban đầu. Nếu bỏ qua đối số **pv**, thì nó được mặc định là 0.

* **Type**: đối số tùy chọn. Số 0 hoặc 1 chỉ rõ thời điểm thanh toán đến hạn. 1 nếu số tiền trả đầu kỳ, 0 nếu số tiền trả cuối kỳ. Nếu đối số kiểu bị bỏ qua, thì nó được mặc định là 0.



### 2. Cách sử dụng hàm FV


Ví dụ: Bạn muốn tiết kiệm một số tiền trong vòng 10 năm. Đầu tiên bạn gửi 100 triệu vào tài khoản tiết kiệm với lãi suất 5%/năm. Cứ mỗi năm tiếp theo bạn gửi vào tài khoản 50 triệu. Hãy tính xem trong tài khoản của bạn sẽ có bao nhiều tiền sau khi kết thúc 10 năm.


![](https://hddtvn.com/wp-content/uploads/2021/01/GXDfh8S.png)


Lúc này ta sẽ có 2 trường hợp xảy ra là:




* Số tiền gốc gửi từ đầu kỳ trước, số tiền gửi thêm bắt đầu từ kỳ tiếp theo. Khi đó Type = 0 (đầu kỳ tiếp theo là cuối của kỳ trước).

* Số tiền gốc gửi từ đầu kỳ trước, số tiền gửi thêm gửi vào cùng năm đầu tiên với tiền gốc. Khi đó Type = 1.



#### **Trường hợp 1:** Số tiền gốc gửi từ đầu kỳ trước, số tiền gửi thêm bắt đầu từ kỳ tiếp theo


Trường hợp này ta sẽ có đối số Type = 0. Ta có công thức tính giá trị khoản đầu tư sau 10 năm như sau:


**=FV(B5;B6;-B4;-B3;0)**


Trong đó:




* **B5** là lãi suất (5%)

* **B6** là kỳ hạn thanh toán (10 năm)

* **B4** là số tiền gửi hàng năm

* **B3** là số tiền gửi ban đầu

* **0** là do số tiền gửi thêm bắt đầu từ kỳ tiếp theo



![](https://hddtvn.com/wp-content/uploads/2021/01/xfFuL3N.png)


#### Trường hợp 2: Số tiền gốc gửi từ đầu kỳ trước, số tiền gửi thêm gửi vào cùng năm đầu tiên với tiền gốc


Trường hợp này ta sẽ có đối số Type = 1. Ta có công thức tính giá trị khoản đầu tư sau 10 năm như sau:


**=FV(B5;B6;-B4;-B3;1)**


Trong đó:




* **B5** là lãi suất (5%)

* **B6** là kỳ hạn thanh toán (10 năm)

* **B4** là số tiền gửi hàng năm

* **B3** là số tiền gửi ban đầu

* **1** là do số tiền gửi thêm gửi vào cùng năm đầu tiên với tiền gốc



![](https://hddtvn.com/wp-content/uploads/2021/01/THpFGo9.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm FV trong Excel. Chúc các bạn thành công!*


moreHàm FV sẽ tính giá trị tại một thời điểm tương lai nào đó mà…


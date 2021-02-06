Cách sử dụng hàm NPER để tính số kỳ của một khoản đầu tư trong Excel
==========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-nper-de-tinh-so-ky-cu%cc%89a-mo%cc%a3t-khoa%cc%89n-dau-tu-trong-excel/)Hàm NPER được sử dụng rộng rãi trong Excel để tính số kỳ của một khoản đầu tư. Hàm NPER có tác dụng trong nhiều lĩnh vực khác nhau như tài chính, kế toán… 1. Cấu trúc hàm NPER Cú pháp hàm: =NPER(rate; pmt; pv; [fv]; [type]) Trong đó: Rate: đối số bắt buộc, là …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm NPER được sử dụng rộng rãi trong Excel để tính số kỳ của một khoản đầu tư. Hàm NPER có tác dụng trong nhiều lĩnh vực khác nhau như tài chính, kế toán…**


### 1. Cấu trúc hàm NPER


Cú pháp hàm: **=NPER(rate; pmt; pv; [fv]; [type])**


*Trong đó:*




* **Rate**: đối số bắt buộc, là lãi suất theo kỳ hạn.

* **Pmt**: đối số bắt buộc, là khoản thanh toán cho mỗi kỳ.

* **Pv**: đối số bắt buộc, là giá trị hiện tại, hoặc số tiền trả một lần hiện tại đáng giá ngang với một chuỗi các khoản thanh toán tương lai.

* **Fv**: đối số tùy chọn, là giá trị tương lai hay số dư tiền mặt bạn muốn thu được sau khi thực hiện khoản thanh toán cuối cùng.

* **Type**: đối số tùy chọn, là thời điểm thanh toán đến hạn.



*Lưu ý:*




*  **Pmt** không đổi trong suốt kỳ hạn. Đối số **pmt** có chứa tiền gốc và lãi, nhưng không chứa các khoản phí và thuế khác.

* Nếu **fv** được bỏ qua, thì nó được giả định là 0 (ví dụ, giá trị tương lai của khoản vay là 0).

* Nếu đối số **Type** là 0 hoặc bỏ qua thì có nghĩa là thời điểm thanh toán đến hạn ở cuối chu kỳ

* Nếu đối số **Type** là 1 thì có nghĩa là thời điểm thanh toán đến hạn ở đầu chu kỳ



### 2. Cách sử dụng hàm NPER


Ví dụ hiện tại bạn có một khoản đầu tư có giá trị 100.000.000 với lãi suất 11%/năm. Mỗi tháng bạn đầu tư thêm 1.000.000. Bạn cần tính xem mất bao nhiêu lâu để khoản đầu tư có giá trị 1.000.000.000.


![](https://hddtvn.com/wp-content/uploads/2021/01/4SYhceX.png)


Áp dụng cấu trúc hàm bên trên ta có công thức tính số kỳ để khoản đầu tư có giá trị như mong muốn như sau:


=NPER(B3/12;B4;B5;B6;0)


Kết quả ta thu được là sẽ mất 183 tháng để khoản đầu tư đạt giá trị mong muốn 1.000.000.000


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm NPER trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm NPER được sử dụng rộng rãi trong Excel để tính số kỳ của một…


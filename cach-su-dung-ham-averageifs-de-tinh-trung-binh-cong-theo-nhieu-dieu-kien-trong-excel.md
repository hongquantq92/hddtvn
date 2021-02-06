Cách sử dụng hàm AVERAGEIFS để tính trung bình cộng theo nhiều điều kiện trong Excel
====================================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-averageifs-de-tinh-trung-binh-cong-theo-nhieu-dieu-kien-trong-excel/)Trong Excel, hàm AVERAGEIFS để tính giá trị trung bình của những số trong một dải ô đáp ứng được cùng lúc nhiều điều kiện. Bạn sẽ thắc mắc hàm AVERAGEIFS khác với AVERAGEIF như thế nào? Rất đơn giản, AVERAGEIF chỉ tính được giá trị trung bình của các ô đáp ứng 1 điều …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong Excel, hàm AVERAGEIFS để tính giá trị trung bình của những số trong một dải ô đáp ứng được cùng lúc nhiều điều kiện. Bạn sẽ thắc mắc hàm AVERAGEIFS khác với AVERAGEIF như thế nào? Rất đơn giản, AVERAGEIF chỉ tính được giá trị trung bình của các ô đáp ứng 1 điều kiện duy nhất. Nếu có nhiều hơn 1 điều kiện, bạn phải sử dụng kết hợp nó với hàm khác, hoặc đơn giản hơn là sử dụng hàm AVERAGEIFS. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm AVERAGEIFS trong Excel.**


### 1. Cấu trúc hàm AVERAGEIFS


Cú pháp hàm: **=AVERAGEIFS(average\_range; criteria\_range1; criteria1; [criteria\_range2; criteria2]; …)**


*Trong đó:*




* **Average\_range**: đối số bắt buộc, là vùng tính giá trị trung bình. Bao gồm các số hoặc tên, mảng hoặc tham chiếu có chứa số.

* **Criteria\_range1, criteria\_range2,…**:**Criteria\_range1** là đối số bắt buộc, **criteria\_range** tiếp theo là đối số tùy chọn. Là vùng điều kiện.

* **Criteria1, criteria2,…**: **Criteria1** là đối số bắt buộc, các **criteria** tiếp theo là đối số tùy chọn. Là điều kiện.



*Lưu ý:*




* Nếu một ô trong vùng điều kiện trống, hàm AVERAGEIFS sẽ mặc định ô đó bằng 0.

* Nếu **average\_range** là trống hoặc dạng văn bản, hàm AVERAGEIFS sẽ trả về giá trị lỗi #DIV0! .

* Mỗi **criteria\_range** phải có cùng kích cỡ và hình dạng.

* Nếu không có ô nào đáp ứng tất cả các tiêu chí, AVERAGEIFS sẽ trả về giá trị lỗi #DIV/0! .

* Có thể dùng các ký tự đại diện, dấu chấm hỏi (?) và dấu sao (*) trong điều kiện để thay thế cho ký tự.



### 2. Cách sử dụng hàm AVERAGEIFS


Ví dụ ta có bảng điểm thi toán và văn của các học sinh như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/gH28zbp.png)


#### a. Tính trung bình điểm thi văn của những học sinh Nữ có điểm thi toán lớn hơn 5


**=AVERAGEIFS(E2:E11;C2:C11;”Nam”;D2:D11;”>5″)**


Trong đó:




* **E2:E11**là vùng điểm thi văn cần tính trung bình cộng

* **C2:C11**là vùng điều kiện giới tính

* **“Nam”**là điều kiện giới tính là Nam

* **D2:D11**là vùng điều kiện điểm thi toán

* **“>5”**là điều kiện điểm thi toán lớn hơn 5



![](https://hddtvn.com/wp-content/uploads/2021/01/GfWQojR.png)


#### b. Tính điểm thi toán trung bình của những học sinh nữ có tên bắt đầu bằng H


Ta có công thức tính như sau: **=AVERAGEIFS(D2:D11;C2:C11;”Nữ”;B2:B11;”H*”)**


Trong đó:




* **D2:D11** là vùng điểm thi văn cần tính trung bình cộng

* **C2:C11**là vùng điều kiện giới tính

* **“Nwx”** là điều kiện giới tính là Nữ

* **B2:B11** là vùng điều kiện tên

* **“H*”** là điều kiện tên bắt đầu bằng H. Dấu * được thay cho các ký tự



![](https://hddtvn.com/wp-content/uploads/2021/01/vaHwBaX.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn sử dụng hàm AVERAGEIFS trong Excel. Chúc các bạn thành công!*


moreTrong Excel, hàm AVERAGEIFS để tính giá trị trung bình của những số trong một…


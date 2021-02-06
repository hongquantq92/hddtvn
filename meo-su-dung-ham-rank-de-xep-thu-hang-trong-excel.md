Mẹo sử dụng hàm Rank để xếp thứ hạng trong Excel
================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-su-dung-ham-rank-de-xep-thu-hang-trong-excel/)Trong quá trình làm việc trên Excel, nhiều lúc bạn sẽ cần xếp thứ hạng cho những con số. Hàm Rank sẽ giúp bạn thực hiện điều này. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Rank để xếp thứ hạng trong Excel. 1. Cú pháp hàm RANK Cú pháp …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình làm việc trên Excel, nhiều lúc bạn sẽ cần xếp thứ hạng cho những con số. Hàm Rank sẽ giúp bạn thực hiện điều này. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Rank để xếp thứ hạng trong Excel.**


### 1. Cú pháp hàm RANK


Cú pháp hàm: **=RANK(number;ref;[order])**.


*Trong đó:*




* **Number**: là số số cần xếp hạng trong danh sách

* **Ref**: là danh sách các số. Các giá trị không phải là số trong mảng sẽ bị bỏ qua.

* **Order**: thứ tự xếp hạng  

Nếu **Order** là số 0 hoặc không nhập, hàm RANK sẽ sắp xếp theo thứ tự giảm dần  

Nếu **Order** là số 1, hàm RANK sẽ sắp xếp theo thứ tự tăng dần



### 2. Cách sử dụng hàm RANK


#### Ví dụ 1: Ta có bảng tiền hàng của các sản phẩm và cần xếp hạng tiền hàng của các sản phẩm.


![](https://hddtvn.com/wp-content/uploads/2021/01/ya96hIs.png)


Nếu xếp hạng theo thứ tự tiền hàng giảm dần. Ta nhập công thức tại ô G2: **=RANK(F2;$F$2:$F$11;0)**. Sau đó sao chép công thức cho các ô còn lại trong cột G. Ta thu được kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/TSIxIrG.png)


Nếu xếp hạng theo thứ tự tiền hàng tăng dần. Ta nhập công thức tại ô G2: **=RANK(F2;$F$2:$F$11;1)**. Sau đó sao chép công thức cho các ô còn lại trong cột G. Ta thu được kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/9NsEp6s.png)


#### Ví dụ 2: Ta có điểm thi môn toán của từng học sinh và cần xếp hạng điểm của từng người theo thứ tự giảm dần.


![](https://hddtvn.com/wp-content/uploads/2021/01/UcekqYR.png)


Ta nhập công thức tại ô D2: **=RANK(C2;$C$2:$C$9;0)**. Sau đó sao chép công thức cho các ô còn lại trong cột G. Ta thu được kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/HL3QiEv.png)


Theo bảng trên, ta thấy Tú, Thành và Hương cùng xếp hạng 1 vì đều đạt 9 điểm. Huyền và Trung đều xếp hạng 6 vì cùng đạt 6 điểm.


Nếu không muốn xếp cùng hạng, ta có thể kết hợp thêm hàm COUNTIF vào công thức để xếp hạng được liên tục mà không có thứ hạng trùng.


Nhập công thức tại ô D2: **=RANK(C2;$C$2:$C$9;0)+COUNTIF($C$2:C2;C2)-1**. Sau đó sao chép công thức cho các ô còn lại trong cột G. Ta thu được kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/HzfWFyE.png)


Lúc này sẽ không còn các thứ hạng trùng. Nếu các số giống nhau thì số nào xuất hiện trước sẽ được thứ hạng trước, số nào xuất hiện sau sẽ được thứ hạng sau.


*Như vậy, bài viết trên đã hướng dẫn cách bạn sử dụng hàm RANK để xếp thứ hạng trong Excel. Chúc các bạn thành công!*


moreTrong quá trình làm việc trên Excel, nhiều lúc bạn sẽ cần xếp thứ hạng…


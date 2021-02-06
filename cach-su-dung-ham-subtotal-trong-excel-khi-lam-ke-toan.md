Cách sử dụng hàm Subtotal trong Excel khi làm kế toán
=====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-subtotal-trong-excel-khi-lam-ke-toan/)Hướng dẫn cách sử dụng hàm Subtotal trong Excel: Hàm tính tổng của các tổng trong Excel kế toán như tính tổng phát sinh trong kỳ, tính tổng cho từng tài khoản… ——————————————————————————————- 1. Công dụng của hàm Subtotal trong Excel đối với kế toán: – Tính tổng dãy ô có điều kiện như:                  …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



Hướng dẫn cách sử dụng hàm Subtotal trong Excel: Hàm tính tổng của các tổng trong Excel kế toán như tính tổng phát sinh trong kỳ, tính tổng cho từng tài khoản…
-----------------------------------------------------------------------------------------------------------------------------------------------------------------



  

![sử dụng hàm subtotal trong excel kế toán](https://hddtvn.com/wp-content/uploads/2021/01/su-dung-ham-subtotal-trong-excel-ke-toan.png "sử dụng hàm subtotal trong excel kế toán")
 ——————————————————————————————-



**1. Công dụng của hàm Subtotal trong Excel đối với kế toán:**


– Tính tổng dãy ô có điều kiện như:  

                 Tính tổng phát sinh trong kỳ.  

                 Tính tổng cho từng tài khoản cấp 1.  

                 Tính tổng tiền tồn cuối ngày.



 ——————————————————  

  
**2. Tác dụng hàm SUBTOTAL:**  

   

– Hàm Subtotal là hàm tính toán cho một nhóm con trong một danh sách hoặc bảng dữ liệu tuỳ theo phép tính mà bạn chọn lựa trong đối số thứ nhất.



**Cú pháp: =SUBTOTAL(function\_num,ref1,ref2,…)**

**Các tham số:**  

– **Function\_num**: là các con số từ 1 đến 11 (hay có thêm 101 đến 111 trong phiên bản Excel 2003, 2007) qui định hàm nào sẽ được dùng để tính toán trong subtotal  

– **Ref1, ref2,..**. là các vùng địa chỉ tham chiếu mà bạn muốn thực hiện phép tính trên đó.  

   

– Đối số thứ nhất của hàm SUBTOTAL bắt buộc bạn phải nhớ con số đại diện cho phép tính cần thực hiện trên tập số liệu. Đối số đó được xác định hàm thực sự nào sẽ được sử dụng khi tính toán trong danh sách bên dưới.



![hướng dẫn sử dụng hàm subtotal trong excel](https://hddtvn.com/wp-content/uploads/2021/01/huong-dan-su-dung-ham-subtotal-trong-excel.jpg "hướng dẫn sử dụng hàm subtotal trong excel")
   

**Ví dụ:**  

– Nếu đối số là 1 thì hàm SUBTOTAL hoạt động giống như hàm AVERAGE, nếu đối số thứ nhất là 9 thì hàm hàm SUBTOTAL hoạt động giống như hàm SUM.


**Chú ý:** Tuy có nhiều đối số là từ 1 – 11 nhưng trong kế toán chúng ta thường sử dụng **đối số 9** và thường sử dụng trong việc tính tổng cho từng tài khoản, tính tổng phát sinh bên Nợ, Có, tỉnh tổng số tiền cuối ngày.


– Cú pháp hàm: 

**= SUBTOTAL(9; dãy ô cần tính tổng)** ( *Số 9 là cú pháp mặc định của hàm cho việc tính tổng).*
**Ví dụ 1**: Muốn cộng Tổng số tiền thu được trong tháng trên sổ Qũy tiền mặt, ta dùng hàm subtotal như sau:  

– Bấm chuột trái -> Sau đó cú phám: **=Subtotal(****9****;G14:G442)**  

*(G14:G442) -> Là dãy ô cần tính tổng, chi tiết như hình ảnh bên dưới*


![cách sử dụng hàm subtotal trong excel](https://hddtvn.com/wp-content/uploads/2021/01/cach-su-dung-ham-subtotal-trong-excel.png "cách sử dụng hàm subtotal trong excel") 



———————————————————————

**Ví dụ 2:** Muốn cộng Tổng số phát sinh bên Nợ trên sổ Nhật ký chung:


![hàm subtotal trong excel](https://hddtvn.com/wp-content/uploads/2021/01/ham-subtotal-trong-excel.png "hàm subtotal trong excel")



  

———————————————————————–

  

**Ví Dụ 3:** Muốn tính được số tiền tồn cuối ngày:  

 – Số tồn ngày hôm trước: **$I$14**  

– Tổng số bên Cột Thu:  **= Subtotal****(9****;****$G$15:G15)**  

– Tổng số bên Cột Chi: **= Subtotal****(9****;****$H$15:H15)**
Công thức tính số tồn cuối ngày: = Tồn ngày hôm trước + Thu – Chi:  

-> Cột tồn tiền cuối ngày dùng hàm Subtotal => Cú pháp hàm:



**=$I$14****+****SUBTOTAL(9****;****$G$15:G15)****–****SUBTOTAL(9****;****$H$15:H15)**

**Cách thực hiện chi tiết**:  

– Sau khi bấm vào Ô số tồn ngày hôm trước (bên trên) tức là ô: =**I14** -> Các bạn bấm phím**F4** (1 lần) để cố định -> Kết quả sẽ là: =**$I$14**


– Tiếp đó sau khi gõ hàm subtotal và bấm vào ô G15 (tức là bên Cột Thu): **=$I$14+SUBTOTAL(9;G15** -> Các bạn cũng bấm phím F4 (1 lần) để có định ô **G15**.  -> kết quả: **=$I$14+SUBTOTAL****(9****;****$G$15**-> Tiếp đó các bạn bấm tiếp 1 lần nữa vào ô G15 -> Kết quả: **=$I$14+SUBTOTAL****(9****;****$G$15:G15)**


– Tiếp theo Hàm subtotal bên cột Chi các bạn cũng làm như vậy nhé, công thức cuối cùng phải như sau:



**=$I$14****+****SUBTOTAL(9****;****$G$15:G15)****–****SUBTOTAL(9****;****$H$15:H15)**

 ![cách sử dụng hàm subtotal](https://hddtvn.com/wp-content/uploads/2021/01/cach-su-dung-ham-subtotal.png "cách sử dụng hàm subtotal")



 ——————————————————  

  
**Ghi chú:**  

– Nếu có hàm subtotal khác lồng đặt tại các đối số ref1, ref2,… thì các hàm lồng này sẽ bị bỏ qua không được tính nhằm tránh trường hợp tính toán 2 lần.  

– Đối số function\_num nếu từ 1 đến 11 thì hàm SUBTOTAL tính toán bao gồm cả các giá trị ẩn trong tập số liệu (hàng ẩn).  

– Đối số function\_num nếu từ 101 đến 111 thì hàm SUBTOTAL chỉ tính toán cho các giá trị không ẩn trong tập số liệu (bỏ qua các giá trị ẩn).



  

 ——————————————————————————–
  


Để có thể làm tốt công việc kế toán trên Excel,các bạn xem thêm: [**Cách làm sổ sách kế toán trên Excel**](# "cách làm sổ sách kế toán trên excel")




moreHướng dẫn cách sử dụng hàm Subtotal trong Excel kế toán: Hàm tính tổng của các tổng trong Excel như tính tổng phát sinh trong kỳ, tính tổng cho từng t…


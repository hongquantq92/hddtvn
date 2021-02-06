Excel: 3 Cách chuyển đổi thời gian thành số thập phân và ngược lại
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/excel-3-cach-chuyen-doi-thoi-gian-thanh-so-thap-phan-va-nguoc-lai/)Nhiều lúc bạn sẽ phải chuyển đổi thời gian thành số thập phân (và ngược lại) để phục vụ cho các thao tác tính toán trong Excel. Để đổi thời gian ra dạng giờ, phút và giây bạn có thể sử dụng phép toán chuyển giá trị thời gian thành số thập phân, hoặc dùng …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Nhiều lúc bạn sẽ phải chuyển đổi thời gian thành số thập phân (và ngược lại) để phục vụ cho các thao tác tính toán trong Excel. Để đổi thời gian ra dạng giờ, phút và giây bạn có thể sử dụng phép toán chuyển giá trị thời gian thành số thập phân, hoặc dùng hàm phù hợp để chuyển đổi các giá trị. Hãy theo dõi bài viết sau để biết cách làm chi tiết nhé.**


### 1. Sử dụng phép toán


Cách đơn giản nhất để chuyển đổi thời gian thành số thập phân trong Excel là nhân giá trị thời gian gốc với số giờ, phút, giây trong ngày:




* Để chuyển đổi thời gian sang giờ, nhân số thời gian với 24 (số giờ trong một ngày).

* Để chuyển đổi thời gian sang phút, nhân với thời gian với 1440 (số phút trong một ngày = 24 * 60).

* Để chuyển đổi thời gian sang giây, nhân với thời gian với 86400 (số giây trong một ngày = 24 * 60 * 60).



Ví dụ ta có bảng thời gian sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/KqGfeRB.png)


Ở cột B tính giờ bạn sử dụng công thức B2=A2 *24, sao chép cho các ô còn lại trong cột B, bạn sẽ thu được kết quả số giờ dưới dạng thập phân.


[![](https://hddtvn.com/wp-content/uploads/2021/01/dKIANFK.png)](https://hddtvn.com/wp-content/uploads/2021/01/dKIANFK.png)


Tương tự, đổi ra phút theo công thức: C2 =A2*24*60


![](https://hddtvn.com/wp-content/uploads/2021/01/ys2vuBi.png)


Đổi ra giây theo công thức: D2 =A2*24*60*60


![](https://hddtvn.com/wp-content/uploads/2021/01/WhJLLhT.png)


### 2. Sử dụng hàm CONVERT


Một cách khác để thực hiện chuyển đổi thời gian sang số thập phân là sử dụng hàm **CONVERT**


Cú pháp hàm: **=CONVERT(number;from\_unit;to\_unit)**


*Trong đó:*




* **Number** là số gốc cần chuyển đổi.

* **from\_unit** là đơn vị của số gốc.

* **to\_unit** là đơn vị cần chuyển đổi.



*Các giá trị của đơn vị:*




* “day” = nếu đơn vị là ngày.

* “hr” = đơn vị là giờ.

* “mn” = đơn vị là phút.

* “sec” = đơn vị là giây.



Với ví dụ trên, giá trị from\_unit mặc định là “day”, và bạn muốn chuyển sang giờ ta có công thức:


**=CONVERT(A2;”day”;”hr”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/kiJfrfQ.png)


Muốn chuyển sang phút ta có công thức:


**=CONVERT(A2;”day”;”mn”)** hoặc: **=CONVERT(B2;”hr”;”mn”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/Ah7haUS.png)


Muốn chuyển sang giây ta có công thức:


**=CONVERT(A2;”day”;”sec”)** hoặc **=CONVERT(B2;”hr”;”sec”)** hoặc **=CONVERT(C2;”mn”;”sec”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/Gtr0Xql.png)


### 3. Sử dụng các hàm HOUR, MINUTE, SECOND


Cuối cùng, bạn có thể sử dụng công thức phức tạp hơn. Xuất các đơn vị thời gian bằng các hàm thời gian HOUR, MINUTE, SECOND sau đó dùng phép toán cộng lại.


Cú pháp hàm:


**= HOUR(serial\_number)**  

**= MINUTE(serial\_number)**  

**= SECOND(serial\_number)**


*Trong đó:* **serial\_number** là số cần chuyển đổi.


Với ví dụ như trên, bạn gõ công thức lần lượt như sau:


**B2 = HOUR (A2)**  

**C2 = MINUTE (A2)**  

**D2 = SECOND (A2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/uyZKGXu.png)


Để tính tổng cộng số giờ ở cột E, công thức tính sẽ bằng số giờ + số phút/60 + số giây/360


Công thức : **=B2+C2/60+D2/3600**. Sao chép công thức cho các dòng còn lại trong cột. Bạn thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/nrcktMZ.png)


Để tính tổng cộng số phút ở cột F, công thức tính sẽ bằng số giờ*60 + số phút + số giây/60


Công thức: **=B2*60+C2+D2/60**. Sao chép công thức cho các dòng còn lại trong cột. Bạn thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/C6dn7u7.png)


Để tính tổng cộng số giây ở cột G, công thức tính sẽ bằng số giờ*3600 + số phút*60 + số giây


Công thức: **=B2*3600+C2*60+D2**. Sao chép công thức cho các dòng còn lại trong cột. Bạn thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/k2GqeLa.png)


*Như vậy, bài viết trên đã hướng dẫn bạn các cách để chuyển đổi thời gian thành giờ, phút giây trong Excel. Chúc các bạn thành công!*


moreNhiều lúc bạn sẽ phải chuyển đổi thời gian thành số thập phân (và ngược…


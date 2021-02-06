Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/bo-tui-cac-ham-dung-de-lam-so-sach-ke-toan-trong-excel/)Một trong những việc mà người kế toán phải làm chính là xây dựng hệ thống sổ sách. Để làm tốt sổ sách kế toán, các bạn cần phải sử dụng rất nhiều đến các hàm trong Excel. Dưới đây là những hàm sẽ hỗ trợ đắc lực cho dân kế toán làm sổ sách …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Một trong những việc mà người kế toán phải làm chính là xây dựng hệ thống sổ sách. Để làm tốt sổ sách kế toán, các bạn cần phải sử dụng rất nhiều đến các hàm trong Excel. Dưới đây là những hàm sẽ hỗ trợ đắc lực cho dân kế toán làm sổ sách một cách chính xác và khoa học.**


### Làm sổ sách kế toán bằng hàm IF


Hàm IF là hàm tính toán có điều kiện, các bạn có thể dùng các dữ liệu cơ sở làm giả thiết để đưa ra kết quả. Đây là hàm thường dùng để tính tiền lương, tiền phòng, tiền điện, nước,… rất cần cho sổ sách của các công ty lớn.


Công thức sử dụng hàm IF như sau:


=IF( Điều kiện, giá trị 1, giá trị 2). Nếu như điều kiện đưa ra ban đầu đúng thì hàm cho kết quả là giá trị 1, còn nếu điều kiện sai thì kết quả trả về sẽ là giá trị 2.  

Ví dụ: Tính tiền giảm trừ của khách thuê phòng dựa vào số ngày lưu trú (số ngày ở). Nếu khách lưu trú từ 20 ngày trở lên giảm 1.000.000 đồng. Từ 10 – 20 ngày giảm 500.000 đồng, còn lại không giảm. Các bạn sử dụng hàm IF như hình dưới đây.


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/91329658_2532073123563877_7955824486547193856_n.png?_nc_cat=104&_nc_sid=b96e70&_nc_eui2=AeGKUXx_Huzk_npQ1B5dvoPNymcnK8ZfXkh4aKZA8E41PCEt3WgnwVI1CyzF8l0wazyoyPT3zZhln4BehiMDZLBH9Jq9-3GZ2I_fmOqPMmU_zA&_nc_ohc=wGJHqS4SiY0AX9HxN_9&_nc_ht=scontent.fsgn2-2.fna&oh=42da0ec79f895403c3f4f2b9b09e4d89&oe=5EA248B2)


[**>>>Cách dùng hàm IF và cách kết hợp với các hàm khác trong Excel**](#)


### Dùng hàm SUM để tính tổng


Hàm SUM là hàm dùng để tính tổng các giá trị, không thể thiếu trong việc làm sổ sách. Cách thức sử dụng hàm này cũng khá đơn giản với cú pháp như sau:


= SUM(số1, số2,…, số hoặc vùng dữ liệu)  

Ví dụ: Nếu các bạn muốn tính tổng của dãy số: 5,10,15,20,25 thì dùng hàm SUM như sau:  

=SUM(5,10,15,20,25) => Kết quả =75


Hàm này thường dùng để tính tổng thu nhập của các nhân viên bao gồm tiền lương và tiền phụ cấp như sau:


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-4.fna.fbcdn.net/v/t1.15752-9/90999315_299976554317709_4984432279993450496_n.png?_nc_cat=106&_nc_sid=b96e70&_nc_eui2=AeGckkJyb3gqDOewqnC5bOSgj3jjTtOTVTymj80QBPV1-QLA3_mSRtP-UfthkXwkJBVg0bdiY4j452KLi0g7NpxUPW1xZj_9o-yJrJmDW-hvjA&_nc_ohc=dF17bYMD240AX-exYKX&_nc_ht=scontent.fsgn2-4.fna&oh=69877343cbe4086fc41e752d33369dbe&oe=5EA1EA18)


### Làm sổ sách kế toán bằng hàm VLOOKUP


Nếu các bạn muốn tìm kiếm hoặc tham chiếu dữ liệu thì chắc chắn không thể không dùng tới hàm VLOOKUP. Cú pháp sử dụng của hàm này như sau:


= VLOOKUP(a, bảng tham chiếu, cột lấy giá trị, kiểu giá trị)


Có thể hiểu đơn giản công thức này là lấy một giá trị (a) đem so sánh theo cột của bảng tham chiếu để trả về giá trị ở cột tương ứng. Kiểu giá trị trong hàm VLOOKUP thường lấy là FALSE.


Chú ý khi chọn vùng giá trị thì các bạn luôn phải lấy giá trị tuyệt đối bằng cách nhấn F4. Cách dùng phổ biến nhất của hàm này là gán đơn giá của hàng hóa như trong ví dụ sau đây.


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/90736595_862116657594430_7483836417127219200_n.png?_nc_cat=100&_nc_sid=b96e70&_nc_eui2=AeEM4nU92nlxLJJvXUtXsuI11Rr-19pLWLq7eB2Mtio22ohDtUl9e1FX1J1TuRg08dkNnat0CHEEr4KMy4pSX0s8jrsjQRTRfNk0BdV41eYTYQ&_nc_ohc=XUwVOTz3vEIAX_sxu4t&_nc_ht=scontent.fsgn2-2.fna&oh=2d06931ff9d3e9078e4b3b42474523b4&oe=5EA47377)


Gán đơn giá cho các mặt hàng sữa Abbolt: 178.000 đồng; Meji: 40.000 đồng; Dutch Lady: 40.000 đồng và Vina Milk: 90.000 đồng. Các bạn sử dụng hàm VLOOKUP theo công thức sẽ ra kết quả ấn định giá của từng sản phẩm như vậy.


[**>>>Hướng dẫn dùng hàm VLOOKUP trong Excel**](#)


### Dùng hàm MIN và MAX để tính giá trị nhỏ nhất, lớn nhất


Hàm MIN và hàm MAX dùng để tính giá trị nhỏ nhất và lớn nhất của vùng dữ liệu. Khi làm sổ sách kế toán, chắc chắn các bạn phải sử dụng đến hai hàm này để cho ra kết quả thống kê chính xác. Công thức của hai hàm này cũng khá đơn giản và dễ hiểu, cụ thể như sau:


= MIN(number 1, number 2,…)  

Ví dụ: Trong dãy số 15,25,35 các bạn muốn tìm giá trị nhỏ nhất thì lập tức có thể sử dụng hàm MIN và kết quả cho ra là số 15.  

MIN(15,25,35) =15


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-4.fna.fbcdn.net/v/t1.15752-9/90578300_688651291677161_2077118687688523776_n.png?_nc_cat=111&_nc_sid=b96e70&_nc_eui2=AeF0dvteLwo1z2rKr7pUXCOAPqMT6lIaUaFIL94bkkYdOuNvex5p381qqwaR0-wT-sovf3jLmIZzwUbbu-FVf6YeSKjCKjnFtn6R9HwlaP1oMg&_nc_ohc=DbVENm_ycJ0AX9nBBRZ&_nc_ht=scontent.fsgn2-4.fna&oh=e58d55e7a73f7143b332d65f70097d77&oe=5EA50105)


Hàm này thường được ứng dụng để tìm ra nhân viên có số ngày nghỉ thấp nhất. Trong cả nghìn nhân viên của một doanh kế toán không thể ngồi đếm số ngày nghỉ của từng người để so sánh được mà hàm MIN sẽ thay bạn làm việc đó.


Ngược lại với hàm MIN, hàm MAX được dùng để tìm ra giá trị lớn nhất với cú pháp:  

= MAX(number 1, number 2, …)  

Ví dụ: Tìm giá trị lớn nhất của dãy số: 100, 150, 250 thì nhập công thức  

MAX(100,150,250), kết quả sẽ ra là 250


[**>>> Cách tìm giá trị lớn nhất và nhỏ nhất bằng hàm MAX và MIN trong Excel**](#)


### Làm sổ sách kế toán bằng hàm LEFT và RIGHT


Hai hàm này thường được dùng để tách các ký tự trái và phải. Thông thường khi nhập hàng hóa, các doanh nghiệp thường dùng mã để dễ nhận biết sản phẩm. Các mã này có thể được phân biệt với nhau bằng 2 chữ cái đầu hoặc 2 chữ cái cuối nên cần dùng hàm LEFT và RIGHT để tách khi làm sổ sách kế toán.


Cú pháp sử dụng của hai hàm này như sau:


= LEFT(chuỗi, ký tự muốn lấy)  

= RIGHT(chuỗi, ký tự muốn lấy)  

Ví dụ: Muốn tách 2 ký tự đầu của sữa Abbolt các bạn sử dụng hàm LEFT sẽ ra kết quả là Ab như hình dưới đây.


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/90554546_221747708881349_737548070585106432_n.png?_nc_cat=102&_nc_sid=b96e70&_nc_eui2=AeHbd3fG4iElHUt05rHZpX8aQn_7c7mwTM-u_7ic0SxH4DUVRDyrvph1oHRNMICpvl61sSmqerjZitrTU7pTCkUfcxV1L85ovrsliaDAMOukTw&_nc_ohc=6fYqj88BW64AX9y7_Ob&_nc_ht=scontent.fsgn2-2.fna&oh=96e231c96d7a202cd50cf71bef24696b&oe=5EA4B011)


**>>>[Cách sử dụng hàm LEFT để trích xuất các ký tự từ chuỗi văn bản](#)**


**>>>**[**Cách sử dụng hàm RIGHT để trích xuất các ký tự từ chuỗi văn bản**](#)


### Dùng hàm AVERAGE để tính trung bình cộng


Đây là hàm dùng để tính giá trị trung bình, cũng là một khâu quan trọng khi làm sổ sách kế toán. Vì thế, người kế toán viên cần biết sử dụng thành thạo hàm này. Cú pháp của hàm AVERAGE như sau:  

=AVERAGE(number 1, number 2,…)


Ví dụ các bạn muốn tính trung bình tiền lương cơ bản thì có thể sử dụng hàm AVERAGE để tính toán như bảng sau:


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-4.fna.fbcdn.net/v/t1.15752-9/90849225_510440249834470_7548092602503921664_n.png?_nc_cat=111&_nc_sid=b96e70&_nc_eui2=AeHjkU7_rBqs-CB3FNLEpJRIS_UyjZbsmI_jX9gWYlKww9BmV_JoXRLlH7sZEsNQbllreWlOPKUw7JoC4ukNJurdvObqy8cOyf2--wuwA1QC7Q&_nc_ohc=0QVRkA1B55sAX-__oGu&_nc_ht=scontent.fsgn2-4.fna&oh=ab417dcc7a323e951c53262453e09967&oe=5EA4A42A)


[**>>> Cách sử dụng hàm Average để tính trung bình cộng trong Excel**](#)


### Dùng bằng hàm DSUM để tính tổng có điều kiện


Nếu các bạn muốn tính tổng số tiền bán được của nhiều mặt hàng khác nhau thì không thể bỏ qua hàm DSUM. Đây là hàm nâng cao của hàm SUM sử dụng nhiều điều kiện hơn.  

Cú pháp sử dụng hàm DSUM:


= DSUM(database, field, criteria)  

Trong đó:  

• Database là toàn bộ cơ sở dữ liệu  

• Field là những mặt hàng cần tính toán  

• Criteria là loại giá trị cần tính


Ví dụ: Cho bảng dữ liệu như hình sau, nếu các bạn muốn tính tổng thành tiền của hai mặt hàng là đĩa mềm và đĩa CD thì sử dụng hàm DSUM như sau:


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/91297982_224606182235808_1196328729022824448_n.png?_nc_cat=107&_nc_sid=b96e70&_nc_eui2=AeFCcYwQ_mbKTbaZny2AiFkbJjvVlmPxYolwQPAXtL_GkO6FPhWmK04xEP4wMKSwPBFFJjyb6quVmb4-AleUGM6Rpt3Lsoka7Bx17yHILoGT2A&_nc_ohc=Xytt6nlEoRYAX-uiWTz&_nc_ht=scontent.fsgn2-2.fna&oh=4f07147a5cb031fee603b5e7092bc140&oe=5EA4D675)


### Làm sổ sách kế toán bằng hàm DAVERAGE


Đây là hàm tính trung bình có điều kiện, nâng cao hơn hàm AVERAGE, người làm kế toán cũng cần đặc biệt chú ý tới hàm này.


Cú pháp của hàm DAVERAGE như sau:


=DAVERAGE(database, field, criteria) tương tự như hàm DSUM


Các bạn có thể ứng dụng hàm DAVERAGE để tính trung bình các mặt hàng bán được nằm trong các khoảng số lượng nhất định. Ví dụ trong bảng dữ liệu sau đây, để tính trung bình số tiền bán được của các mặt hàng có số lượng từ 5 đến 8 các bạn làm như sau:


![Bỏ túi các hàm dùng để làm sổ sách kế toán trong Excel](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/90554546_1120397244971478_2009157285565169664_n.png?_nc_cat=104&_nc_sid=b96e70&_nc_eui2=AeGKOS25-Z0raO1ZKJalrj6yi5jxmh322h3oOV5ObID6KkUwxLqYKXR58muq2YBJqHLQKnxW3QgJZLBC1DmtEokAhuVbn4DZ74Hw-xW6q-j03Q&_nc_ohc=58C7JHmQzxAAX_KRkPL&_nc_ht=scontent.fsgn2-2.fna&oh=f9d3f5b96e9f710a4aac5195ab7bb710&oe=5EA4E268)


Chú ý: Với các hàm có điều kiện thì các bạn cần lập một bảng phụ điều kiện bên cạnh sẽ trực quan hơn khi làm sổ sách kế toán.


Chúc các bạn sử dụng thành công những hàm trên đây để thực hiện tốt công việc làm sổ sách kế toán của mình.



moreMột trong những việc mà người kế toán phải làm chính là xây dựng hệ…


Cách tách thương và số dư bằng hàm QUOTIENT và hàm MOD trong Excel
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-tach-thuong-va-so-du-bang-ham-quotient-va-ham-mod-trong-excel/)Khi làm việc với Excel các bạn cần thực hiện các phép tính toán, trong đó có phép tính chia. Với các phép chia có dư, chúng ta cần sử dụng hàm QUOTIENT và hàm MOD để tách thương và số dư ra hai cột khác nhau. Bài viết sau đây sẽ hướng dẫn các …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi làm việc với Excel các bạn cần thực hiện các phép tính toán, trong đó có phép tính chia. Với các phép chia có dư, chúng ta cần sử dụng hàm QUOTIENT và hàm MOD để tách thương và số dư ra hai cột khác nhau. Bài viết sau đây sẽ hướng dẫn các bạn sử dụng hai hàm này để tách thương và số dư khi thực hiện phép chia trong Excel.**


### Tại sao cần sử dụng kết hợp hàm QUOTIENT và hàm MOD trong Excel?


Nếu bạn muốn tách thương và số dư của cùng một phép tính thì bắt buộc phải sử dụng kết hợp hàm QUOTIENT và hàm MOD. Bởi vì một hàm dùng để tách thương còn một hàm dùng để tách số dư. Hai hàm này luôn đi kèm với nhau, nếu bỏ một trong hai thì thao tác tách thương và số dư sẽ không thể thực hiện được.


### Cách dùng hàm QUOTIENT để tách thương trong Excel


Khi thực hiện phép chia có dư như 10 chia 3 chúng ta sẽ được thương là 3, dư 1. Nếu dùng hàm QUOTIENT kết quả sẽ là số nguyên là 3. Cú pháp của hàm này như sau:  

=QUOTIENT(numerator, denominator)  

Trong đó:  

• Numerator là số bị chia  

• Denominator là số chia  

Ví dụ: Thực hiện các phép tính trong bảng sau bằng hàm QUOTIENT


![](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/90989308_243439040153104_903209468907487232_n.png?_nc_cat=105&_nc_sid=b96e70&_nc_eui2=AeEYONeVve2ztNKbE1LfZUWmbO0UzEhx1wUs_OPKlX9dXTXAhmoQd3NZ2BgYeBUKujRmqUtM_KethPNdCsY7u13xR_etdca1f5VAo2YUg9N9YA&_nc_ohc=5tQdEJKhjIYAX-BsGO-&_nc_ht=scontent.fsgn2-2.fna&oh=ff99f66ebae6fae37588dc758979029f&oe=5EA545DF)


Tại vị trí C4 các bạn nhập công thức như sau: =QUOTIENT(A4,B4) sẽ ra kết quả là 46. Tương tự như vậy, các bạn coppy công thức xuống các cột tiếp theo sẽ được các kết quả:


![](https://scontent.fsgn2-4.fna.fbcdn.net/v/t1.15752-9/90942056_2536356486682991_6207096631783849984_n.png?_nc_cat=101&_nc_sid=b96e70&_nc_eui2=AeGTX0BcYaCL9WaX_7VfLqpXXT_9-8KP9--UHeOkgFHzGsF6JhwlTz0ZMDnnxT0KG_p8YG9xXii13Srcs0exdm_aCC1W8QmGhM2jIOmxhfQTIg&_nc_ohc=JCN76K-PCjkAX9m9xlZ&_nc_ht=scontent.fsgn2-4.fna&oh=70aa1736f631f0bafb9bc25d55846b6f&oe=5EA1CFEC)


### Cách dùng hàm MOD để tách số dư trong Excel


Trong phép chia có dư ví dụ 10 chia 3 được 3 dư 1, thì khi dùng hàm MOD các bạn sẽ ra kết quả là 1. Cấu trúc của hàm MOD như sau:  

=MOD(number, divisor)  

Trong đó:  

• Number là số bị chia  

• Divisor là số chia  

Ví dụ: Thức hiện các phép tính tách số dư trong bảng trên bằng hàm MOD  

Ở vị trí D4 các bạn nhập công thức =MOD(A4,B4) sẽ ra kết quả là 4 như hình dưới đây.


![](https://scontent.fsgn2-4.fna.fbcdn.net/v/t1.15752-9/90578306_310451869925438_4705988798574493696_n.png?_nc_cat=111&_nc_sid=b96e70&_nc_eui2=AeHcbGXMcee9C0_SPsRyfIWuzMutsQ2KrpgWNCRb4z7KvtYLXVlabkrJftDSoczYMksG_b29IPW008lZNb2X1-miMUa6glC_cpSe-YS0cdMGqw&_nc_ohc=hVbIa1OhtJcAX-Vuq9A&_nc_ht=scontent.fsgn2-4.fna&oh=830be80b60bf1f0faaddd09f71866dc2&oe=5EA3E069)


Copy công thức bằng cách đưa con trỏ chuột sang góc phía dưới bên phải của ô D4 khi xuất hiện dấu + thì kéo xuống đến ô D9 ta được kết quả:


![](https://scontent.fsgn2-2.fna.fbcdn.net/v/t1.15752-9/91526087_2672372799659002_1457118795343593472_n.png?_nc_cat=104&_nc_sid=b96e70&_nc_eui2=AeF7DKGeQh7vJGqRj1w3abIZ4733hGzzdworA-Ql3OfyVSE3paJ8kxRI2r2-T7a2taUgdy4huvm9hsCNHAHJ-1hBKRMNBziXOBTMCbHVTt5YhQ&_nc_ohc=4b0UkZRTnhQAX9cDgoO&_nc_ht=scontent.fsgn2-2.fna&oh=c045dea65baa16e49ddabdf35efc07e5&oe=5EA47E93)


Như vậy là các bạn đã thực hiện tách thành công thương/ số nguyên và số dư ra 2 cột khác nhau trong một phép chia có dư bằng hàm QUOTIENT và hàm MOD. Chỉ với thủ thuật này các bạn sẽ dễ dàng làm việc cùng các số liệu có số dư mà không cần làm tròn số thập phân. Từ đó có thể tổng kết và giữ lại giá trị kế tiếp một cách chính xác nhất.


Chúc các bạn sử dụng thành công hàm QUOTIENT và hàm MOD khi thực hiện phép chia có dư trong Excel!



moreKhi làm việc với Excel các bạn cần thực hiện các phép tính toán, trong…


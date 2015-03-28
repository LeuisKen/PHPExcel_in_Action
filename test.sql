-- 对于数据库设置的一些简要说明 可以结合/App/Conf/config.php
create table star_user(
	id int unsigned not null primary key auto_increment,
	username char(20) not null default '',
	stu_id char(20) not null default '',
	tel char(20) not null default '',
	qq char(20) not null default ''
) engine myisam default charset utf8;
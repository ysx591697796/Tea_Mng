-- MySQL dump 10.13  Distrib 8.0.15, for Win64 (x86_64)
--
-- Host: localhost    Database: tea_mng
-- ------------------------------------------------------
-- Server version	8.0.15

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
 SET NAMES utf8 ;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `class`
--


DROP TABLE IF EXISTS `class`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `class` (
  `cid` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(20) DEFAULT NULL,
  `grade` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`cid`)
) ENGINE=MyISAM AUTO_INCREMENT=4 DEFAULT CHARSET=gbk;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `class`
--

LOCK TABLES `class` WRITE;
/*!40000 ALTER TABLE `class` DISABLE KEYS */;
INSERT INTO `class` VALUES (1,'嵌入一班','17'),(2,'嵌入二班','17'),(3,'移动一班','18');
/*!40000 ALTER TABLE `class` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `grade`
--

DROP TABLE IF EXISTS `grade`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `grade` (
  `gid` int(11) NOT NULL AUTO_INCREMENT,
  `grade` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`gid`)
) ENGINE=MyISAM AUTO_INCREMENT=29 DEFAULT CHARSET=gbk;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `grade`
--

LOCK TABLES `grade` WRITE;
/*!40000 ALTER TABLE `grade` DISABLE KEYS */;
INSERT INTO `grade` VALUES (1,'17'),(2,'18'),(28,'19');
/*!40000 ALTER TABLE `grade` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `stuinfo`
--

DROP TABLE IF EXISTS `stuinfo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `stuinfo` (
  `sid` int(11) NOT NULL AUTO_INCREMENT,
  `username` varchar(20) CHARACTER SET gbk COLLATE gbk_bin NOT NULL,
  `password` varchar(20) DEFAULT NULL,
  `name` varchar(20) DEFAULT NULL,
  `id` varchar(10) DEFAULT NULL,
  `sex` varchar(5) DEFAULT NULL,
  `class` varchar(20) DEFAULT NULL,
  `politics` varchar(20) DEFAULT NULL,
  `phone` varchar(20) DEFAULT NULL,
  `birthday` varchar(40) DEFAULT NULL,
  `status` int(2) DEFAULT '2',
  `grade` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`sid`)
) ENGINE=MyISAM AUTO_INCREMENT=29 DEFAULT CHARSET=gbk;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `stuinfo`
--

LOCK TABLES `stuinfo` WRITE;
/*!40000 ALTER TABLE `stuinfo` DISABLE KEYS */;
INSERT INTO `stuinfo` VALUES (1,'fzy','123456','冯泽宇','171164201','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(2,'lh','123456','罗欢','171164202','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(3,'pl','123456','潘蕾','171164203','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(4,'zjn','123456','张佳宁','171164205','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(5,'ztt','123456','张婷婷','171164206','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(6,'zyc','123456','周盈晨','171164207','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(7,'xkl','123456','肖凯莉','171164208','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(8,'mxy','123456','马歆怡','171164209','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(9,'cr','123456','陈瑞','171164210','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(10,'fhj','123456','冯慧娟','171164211','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(11,'wby','123456','王碧媛','171164212','女','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(12,'yds','123456','杨德帅','171164225','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(13,'zgx','123456','赵国兴','171164226','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(14,'lyz','123456','刘雨舟','171164227','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(15,'xhf','123456','徐鸿飞','171164228','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(16,'whd','123456','王宏德','171164229','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(17,'lhq','123456','李和桥','171164230','男','嵌入一班','共青团员','12345678911','2013-08-22T05:03:00.154Z',2,'17'),(18,'sz','123456','申钊','171164231','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(19,'hcl','123456','贺成龙','171164232','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(20,'ljz','123456','梁家桢','171164233','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(21,'nz','123456','牛忠','171164234','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(22,'lzh','123456','李泽辉','171164235','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(23,'lxh','123456','李晓辉','171164236','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(24,'zlb','123456','赵良彬','171164237','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(25,'zy','123456','张义','171164238','男','嵌入一班','共青团员','199','1999-06-20T14:39:50.864Z',2,'17'),(26,'ysx','123456','杨圣轩','171164239','男','嵌入一班','共青团员','15836159','2019-12-17T14:25:20.610Z',2,'17'),(27,'lrq','123456','逯瑞琪','171164240','男','嵌入一班','共青团员','123456','0000-0000-0000',2,'17'),(28,'root','123456','root','管理员','男','管理员','管理员','123456789','管理员',1,'管理员');
/*!40000 ALTER TABLE `stuinfo` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `vacaterecord`
--

DROP TABLE IF EXISTS `vacaterecord`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `vacaterecord` (
  `vid` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(20) DEFAULT NULL,
  `id` varchar(10) DEFAULT NULL,
  `class` varchar(20) DEFAULT NULL,
  `phone` varchar(20) DEFAULT NULL,
  `grade` varchar(20) DEFAULT NULL,
  `reason` varchar(100) DEFAULT NULL,
  `starttime` varchar(40) DEFAULT NULL,
  `endtime` varchar(40) DEFAULT NULL,
  `status` varchar(10) DEFAULT 'notcheck',
  `type` varchar(20) DEFAULT NULL,
  `result` varchar(20) DEFAULT '未审核',
  PRIMARY KEY (`vid`)
) ENGINE=MyISAM AUTO_INCREMENT=41 DEFAULT CHARSET=gbk;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `vacaterecord`
--

LOCK TABLES `vacaterecord` WRITE;
/*!40000 ALTER TABLE `vacaterecord` DISABLE KEYS */;
INSERT INTO `vacaterecord` VALUES (1,'杨圣轩','171164239','嵌入一班','123456','17','感冒','2019-12-9','2019-12-12','notcheck','病假','同意'),(2,'张义','171164238','嵌入一班','123456','17','家中有事','2019-12-1','2019-12-10','check','事假','驳回'),(28,'赵良彬','171164237','嵌入一班','13703431234','17','测试','2019-12-17','2019-12-20','notcheck','事假','驳回'),(29,'李和桥','171164230','嵌入一班','12345678911','17','测试','2019-12-17T05:17:25.428Z','2019-12-21T05:17:25.428Z','notcheck','病假','同意'),(30,'李和桥','171164230','嵌入一班','1','17','测试123','2019-12-26T16:00:38.944Z','2019-12-27T16:00:38.944Z','notcheck','事假','同意'),(31,'李和桥','171164230','嵌入一班','12345678911','17','测试2','2019-12-25T05:24:07.075Z','2019-12-27T05:24:07.075Z','notcheck','事假','同意'),(32,'李和桥','171164230','嵌入一班','12345678911','17','测试3','2019-12-19T05:27:22.969Z','2019-12-20T05:27:22.969Z','notcheck','病假','同意'),(33,'李和桥','171164230','嵌入一班','12345678911','17','测试1','2019-12-27T16:54:29.966Z','2019-12-28T16:54:29.966Z','notcheck','事假','同意'),(34,'张义','171164238','嵌入一班','15836159167','17','测试','2019-12-18T05:31:55.321Z','2019-12-20T05:31:55.321Z','notcheck','事假','驳回'),(35,'杨圣轩','171164239','嵌入一班','15836159167','17','测试','2019-12-19T13:08:54.887Z','2019-12-26T13:08:54.887Z','notcheck','事假','同意'),(36,'杨圣轩','171164239','嵌入一班','15836159','17','测试驳回2','2019-12-20T16:12:14.055Z','2019-12-21T16:12:14.055Z','notcheck','事假','同意'),(37,'李和桥','171164230','嵌入一班','12345678911','17','333','2019-12-21T14:35:38.513Z','2019-12-28T14:35:38.513Z','notcheck','事假','同意'),(38,'杨圣轩','171164239','嵌入一班','1583615941','17','3333','2019-12-20T14:25:28.551Z','2019-12-21T14:25:28.551Z','notcheck','事假','同意'),(39,'李和桥','171164230','嵌入一班','1232323','17','213','2019-12-21T14:36:01.898Z','2019-12-22T14:36:01.898Z','notcheck','事假','驳回'),(40,'杨圣轩','171164239','嵌入一班','15836159','17','ces','2020-01-23T12:54:11.584Z','2020-01-24T12:54:11.584Z','notcheck','事假','未审核');
/*!40000 ALTER TABLE `vacaterecord` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Dumping events for database 'tea_mng'
--

--
-- Dumping routines for database 'tea_mng'
--
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2019-12-23  0:24:26

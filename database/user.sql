
--
-- Database: `Library`
--

-- --------------------------------------------------------

--
-- Table structure for table user
--

CREATE TABLE IF NOT EXISTS user (
  ID int(15) NOT NULL,
  Rollno int(4) NOT NULL,
  Name varchar(20) NOT NULL,
  Branch varchar(50) NOT NULL,
  PRIMARY KEY (ID)
);

--
-- Dumping data for table `user`
--


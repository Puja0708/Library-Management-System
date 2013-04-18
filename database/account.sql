
--
-- Database: `Library`
--

-- --------------------------------------------------------

--
-- Table structure for table `account`
--

CREATE TABLE IF NOT EXISTS account (
  Acc_no int(15) NOT NULL,
  Total_limit int(3) NOT NULL,
 Books_left text NOT NULL,
 Status text NOT NULL,
  PRIMARY KEY (Acc_no)
);

--
-- Dumping data for table `account`
--


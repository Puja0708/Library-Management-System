-- Database: `Library`
--

-- --------------------------------------------------------

--
-- Table structure for table `book`
--

CREATE TABLE IF NOT EXISTS book (
  B_no int(15) NOT NULL,
  ISBN int(15) NOT NULL,
  Subject text NOT NULL,
  Name varchar(50) NOT NULL,
  Author varchar(50) NOT NULL,
  Publisher varchar(20) NOT NULL,
  Editor varchar(20) NOT NULL,
  Copies int(3) NOT NULL,
  Cost int(4) NOT NULL,
  PRIMARY KEY (B_no)
);

--
-- Dumping data for table `book`
--


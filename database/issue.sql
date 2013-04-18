
--
-- Database: `Library`
--

-- --------------------------------------------------------

--
-- Table structure for table `issue`
--

CREATE TABLE IF NOT EXISTS issue (
  Bno int(11) NOT NULL,
  Id int(11) NOT NULL,
  issue_date date NOT NULL,
  due_date date NOT NULL,
  copies_available int(11) NOT NULL,
  PRIMARY KEY (Id),
  UNIQUE KEY Bno (Bno)
) ;

--
-- Dumping data for table `issue`
--

